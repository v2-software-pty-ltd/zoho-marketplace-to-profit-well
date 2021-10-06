require('dotenv').config()
const fs = require('fs')
const xlsx = require('node-xlsx').default
const axios = require('axios').default
const glob = require('glob-promise')

const PLANS_TO_SYNC = [
    "Advanced Round Robin Lead Assignment",
    "Notes Filter Extension",
    "Stale Lead Tracker",
    "Twilio SMS Extension for Zoho CRM",
    "Smooth Messenger - Twilio SMS for Zoho CRM"
]

const PROFITWELL_ADD_SUBSCRIPTION_ENDPOINT = 'https://api.profitwell.com/v2/subscriptions/'

async function getCustomers() {
    const fileNames = await glob(`${__dirname}/Customers*.xls`)
    console.log(fileNames)

    return fileNames.flatMap((fileName) => {
        const [workSheet] = xlsx.parse(fs.readFileSync(fileName));
        const [headers, ...rows] = workSheet.data

        const customers = rows.map((rowData) => {
            return rowData.reduce((rowAsObject, cellValue, cellIndex) => {
                return {
                    ...rowAsObject,
                    [headers[cellIndex]]: cellValue
                }
            }, {})
        })

        return customers
    })
}

async function getCancellations() {
    const fileNames = await glob(`${__dirname}/cancel*.xls`)

    return fileNames.flatMap((fileName) => {
        const [workSheet] = xlsx.parse(fs.readFileSync(fileName));
        const [headers, ...rows] = workSheet.data

        return rows.map((rowData) => {
            return rowData.reduce((rowAsObject, cellValue, cellIndex) => {
                return {
                    ...rowAsObject,
                    [headers[cellIndex]]: cellValue
                }
            }, {})
        })
    })
}

async function unchurnCustomer(customer) {
    const UNCHURN_URL = `https://api.profitwell.com/v2/unchurn/${customer['Custom Id']}/`
    try {
        const churnResult = await axios.put(UNCHURN_URL, null, {
            headers: {
                'Authorization': process.env.PROFITWELL_API_KEY
            }
        })
    } catch (e) {
        if (!e.response?.data?.non_field_errors?.[0].includes('was not churned in the fi')) {
            console.error('uh oh unchurn', UNCHURN_URL)
            console.error(customer)
            console.error(e.response?.data)
            // throw e
        }
    }
}

async function wait(seconds) {
    await new Promise((resolve) => {
        setTimeout(resolve, seconds * 1000)
    })
}

async function createSubscription(customer, attempts = 0) {
    const [currency, amount] = customer['Renewal Amount'].split(' ')
    const renewalAmountCents = parseInt(amount, 10) * 100
    const dataForProfitwell = {
        "user_alias": customer['Profile Id'],
        "subscription_alias": customer['Custom Id'],
        "email": customer['Custom Id'],
        "plan_id": customer['Service'],
        "plan_interval": customer['Payperiod'] === 'Yearly' ? 'Year' : 'Month',
        "value": renewalAmountCents,
        // "status": customer['Status'] === 'Inactive' ? 'inactive' : 'active',
        "plan_currency": currency.toLowerCase(),
        "effective_date": new Date(customer['Registration Date']).getTime() / 1000
    }

    try {
        const result = await axios.post(PROFITWELL_ADD_SUBSCRIPTION_ENDPOINT, dataForProfitwell, {
            headers: {
                'Authorization': process.env.PROFITWELL_API_KEY
            }
        })
    } catch (e) {
        if (e.response?.data?.non_field_errors?.[0].includes('already exists')) {
            return
        }
        if (attempts > 3) {
            console.error('uh oh - error with new customer', attempts, e.response?.data)
            // throw new Error('Problem creating')

            return
        }

        console.log('retrying ', customer['Profile Id'], e.response?.data)
        await wait((attempts + 1) * 5)
        console.log('finished waiting for ', customer['Profile Id'])

        await createSubscription(customer, attempts + 1)
    }
}

async function handleCancellation(customer, cancellations, attempts = 0) {
    const cancellationData = cancellations.find((cancellation) => {
        return cancellation['Profile Id'] === customer['Profile Id']
    })

    const cancelType = cancellationData?.['Cancel Type'] ?? 'User Cancel'
    const churnType = cancelType === 'Auto cancel' ? 'delinquent' : 'voluntary'

    const churnDate = new Date(customer['Renewal Date']).getTime() / 1000
    const CHURN_URL = `https://api.profitwell.com/v2/subscriptions/${customer['Custom Id']}/?effective_date=${churnDate}&churn_type=${churnType}`

    try {
        const churnResult = await axios.delete(CHURN_URL, {
            headers: {
                'Authorization': process.env.PROFITWELL_API_KEY
            }
        })
    } catch (e) {
        if (e.response?.data?.non_field_errors?.[0].includes('already scheduled to churn')) {
            return;
        }
        if (!e.response?.data?.non_field_errors?.[0].includes('already scheduled to churn') && attempts > 3) {
            console.error('uh oh churn', CHURN_URL)
            console.error(customer)
            console.error(e.response?.data)

            return
            // throw e
        }
        await wait((attempts + 1) * 15)

        await handleCancellation(customer, cancellations, attempts + 1)
    }
}

async function pushDataToProfitWell() {
    const customers = await getCustomers()
    const cancellations = await getCancellations()

    await Promise.all(customers.slice(0).map(async (customer, customerIdx) => {
        if (!PLANS_TO_SYNC.includes(customer['Service'])) {
            console.log('ignoring', customer['Service'])
            return
        }

        await wait(customerIdx)

        await createSubscription(customer)

        if (customer['Status'] === 'Inactive') {
            await handleCancellation(customer, cancellations)
        } else {
            await unchurnCustomer(customer)
        }

        if (customerIdx % 10 === 0) {
            console.log(`Done with ${customerIdx} of ${customers.length}`)
        }
    }))
}

pushDataToProfitWell()
