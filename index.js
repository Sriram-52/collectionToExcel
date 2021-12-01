var admin = require('firebase-admin')
var xl = require('excel4node')
var addDays = require('date-fns/addDays')

var serviceAccount = require('./admin.json') // path/to/serviceAccountKey.json
var wb = new xl.Workbook()

admin.initializeApp({
	credential: admin.credential.cert(serviceAccount),
})

var db = admin.firestore()

async function paymentsToExcel() {
	try {
		var ws = wb.addWorksheet('Sheet 1')

		const ref = db.collection('PAYMENTS').orderBy('purchasedDate')

		const paymentsSnapshot = await ref.get()

		const payments = paymentsSnapshot.docs.map((doc) => doc.data())

		const groups = payments.reduce((groups, payment) => {
			const date = addDays(payment.purchasedDate.toDate(), 1)
				.toISOString()
				.split('T')[0]
			if (!groups[date]) {
				groups[date] = []
			}
			groups[date].push(payment)
			return groups
		}, {})

		const groupArrays = Object.keys(groups).map((date) => {
			groups[date].sort((a, b) => a.cretedAt.toDate() - b.cretedAt.toDate())
			return {
				date,
				payments: groups[date],
			}
		})

		var i = 2
		groupArrays.forEach((group) => {
			group.payments.forEach((payment) => {
				ws.cell(i, 1).string(group.date)
				ws.cell(i, 2).string(payment.title)
				ws.cell(i, 3).string(payment.description)
				ws.cell(i, 4).number(payment.rate)
				i++
			})
		})
		wb.write('Trip.xlsx')
	} catch (error) {
		console.error(error)
	}
}

paymentsToExcel()
