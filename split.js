const XLSX = require('xlsx')
const workbook = XLSX.readFile('fullroster.xlsx')

function reformatStudents(worksheet) {
	const keys = Object.keys(worksheet)

	let ssns = []
	let ids = []
	let familyNames = []
	let firstNames = []
	let englishNames = []
	let departments = []
	let courseNumbers = []
	let sections = []
	let siasIds = []

	// extract columns
	for (let i=0; i<keys.length; i++) {
		switch (keys[i][0]) {
			case 'A':
				ssns.push(worksheet[keys[i]].v)
				break
			case 'B':
				ids.push(worksheet[keys[i]].v)
				break
			case 'C':
				familyNames.push(worksheet[keys[i]].v)
				break
			case 'D':
				firstNames.push(worksheet[keys[i]].v)
				break
			case 'E':
				englishNames.push(worksheet[keys[i]].v)
				break
			case 'F':
				departments.push(worksheet[keys[i]].v)
				break
			case 'G':
				courseNumbers.push(worksheet[keys[i]].v)
				break
			case 'H':
				sections.push(worksheet[keys[i]].v)
				break
			case 'I':
				siasIds.push(worksheet[keys[i]].v)
				break
		}
	}

	let students = []

	for (let i=0; i<ssns.length; i++) {
		let year = siasIds[i].slice(0, 4)

		let yearOfStudy = 'NA'

		switch (year) {
			case '2016':
				yearOfStudy = 'SP'
				break
			case '2017':
				yearOfStudy = 'JR'
				break
			case '2018':
				yearOfStudy = 'SR'
				break
		}

		let formattedName

		if (englishNames[i] !== '') {
			formattedName = `${familyNames[i]}, ${firstNames[i]} (${englishNames[i]})`
		} else {
			`${familyNames[i]}, ${firstNames[i]}`
		}

		students.push({
			'Student Name': formattedName,
			'Student ID': ids[i],
			'Class': yearOfStudy,
			'Status': 'EN',
			'Status Date': 'NA',
			'Perm. City': 'NA',
			'Last Attended': 'NA',
			'MT Grade': 'NA',
			'Final Grade': 'NA',
			'Grade Change': 'NA',
			'Advisor': 'NA',
			'Department': departments[i],
			'Section': sections[i],
			'Course': courseNumbers[i],
			'Sias ID': siasIds[i]
		})
	}

	return students
}

function buildCourseSections(students) {
	let courseSections = {}

	let columnNames = {
		'Student Name': 'Student Name',
		'Student ID': 'Student ID',
		'Class': 'Class',
		'Status': 'Status',
		'Status Date': 'Status Date',
		'Perm. City': 'Perm. City',
		'Last Attended': 'Last Attended',
		'MT Grade': 'MT Grade',
		'Final Grade': 'Final Grade',
		'Grade Change': 'Grade Change',
		'Advisor': 'Advisor',
		'Department': 'Department',
		'Section': 'Section',
		'Course': 'Course',
		'Sias ID': 'Sias ID'
	}

	for (let i=0; i<students.length; i++) {
		let student = students[i]

		if (courseSections.hasOwnProperty(`${student['Department']} ${student['Course']} ${student['Section']}`)) {
			courseSections[`${student['Department']} ${student['Course']} ${student['Section']}`].push(student)
		} else {
			courseSections[`${student['Department']} ${student['Course']} ${student['Section']}`] = []
			courseSections[`${student['Department']} ${student['Course']} ${student['Section']}`].push(columnNames)
			courseSections[`${student['Department']} ${student['Course']} ${student['Section']}`].push(student)
		}
	}

	return courseSections
}

function saveFile(workbook, filename) {
	XLSX.writeFile(workbook, filename)
}

function split() {
	// find the roster sheet
	const worksheet = workbook.Sheets['Fall 2019 pre-roster']

	// reformat students to more closely match Tiger Central output
	let students = reformatStudents(worksheet)

	// build course sections from reformatted students
	let courseSections = buildCourseSections(students)

	// convert JSON to xls and save
	let sectionKeys = Object.keys(courseSections)
	let sectionValues = Object.values(courseSections)

	for (let i=1; i<sectionKeys.length; i++) {
		let worksheet = XLSX.utils.json_to_sheet(sectionValues[i], {skipHeader: true})

		let book = XLSX.utils.book_new()
		XLSX.utils.book_append_sheet(book, worksheet, 'Course Roster')
		
		saveFile(book, `dest/${sectionKeys[i]}.xls`)
	}
	
}

// do that magic
split()
