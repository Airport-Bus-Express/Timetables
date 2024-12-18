const XLSX = require('xlsx');
const xmlbuilder2 = require('xmlbuilder2');
const fs = require('fs').promises;
const STOPS = require('./timetables/stops.json');

// TransXChange Configuration
const OPERATOR_ID = 'OId_Example';
const SERVICE_CODE = 'PB0000123';
const LINE_NAME = 'XMAS';
const DESCRIPTION = 'Christmas Special Service';

// Define service patterns
const OUTBOUND_STOPS = ['STD', 'STR', 'BETH', 'LIV', 'BETH'];
const RETURN_STOPS = ['LIV', 'BETH', 'STR', 'STD'];

// Function to format time string from Excel format to TXC format (HH:MM)
function formatTimeString(timeStr) {
  if (!timeStr) return null;

  // Convert the decimal time format (e.g., "2.40" or "0.50") to HH:MM format
  const [hours, minutes] = timeStr.split('.');

  // Parse hours and minutes as integers
  let hh = parseInt(hours, 10);
  let mm = minutes ? parseInt(minutes, 10) : 0;

  // Handle cases where minutes are given in decimal format
  // e.g., "2.4" should be interpreted as "2:40" not "2:04"
  if (minutes && minutes.length === 1) {
    mm = mm * 10;
  }

  // Ensure hours and minutes are padded with zeros
  const paddedHours = hh.toString().padStart(2, '0');
  const paddedMinutes = mm.toString().padStart(2, '0');

  return `${paddedHours}:${paddedMinutes}`;
}

async function createTXC(inputFile) {
  try {
    // Read Excel file
    const workbook = XLSX.readFile(inputFile);

    // Create root XML document with namespaces
    const doc = xmlbuilder2.create({ version: '1.0', encoding: 'UTF-8' })
      .ele('TransXChange', {
        'xmlns': 'http://www.transxchange.org.uk/',
        'xmlns:xsi': 'http://www.w3.org/2001/XMLSchema-instance',
        'xsi:schemaLocation': 'http://www.transxchange.org.uk/ http://www.transxchange.org.uk/schema/2.4/TransXChange_general.xsd',
        'CreationDateTime': new Date().toISOString(),
        'ModificationDateTime': new Date().toISOString(),
        'Modification': 'new',
        'RevisionNumber': '1',
        'FileName': 'ChristmasService.xml',
        'SchemaVersion': '2.4'
      });

    // Add Services section
    const services = doc.ele('Services');
    const service = services.ele('Service')
      .ele('ServiceCode').txt(SERVICE_CODE).up()
      .ele('Lines')
      .ele('Line')
      .ele('LineName').txt(LINE_NAME).up()
      .ele('Description').txt(DESCRIPTION).up().up().up();

    // Create operating period
    const operatingPeriod = service
      .ele('OperatingPeriod')
      .ele('StartDate').txt('2024-12-24').up()
      .ele('EndDate').txt('2025-01-01').up().up();

    // Create operating profile
    const operatingProfile = service.ele('OperatingProfile');
    const regularDayType = operatingProfile
      .ele('RegularDayType')
      .ele('DaysOfWeek')
      .ele('Christmas').up()
      .ele('BoxingDay').up()
      .ele('NewYearsEve').up()
      .ele('NewYearsDay').up();

    // Add service organization
    service
      .ele('RegisteredOperatorRef').txt(OPERATOR_ID).up()
      .ele('PublicUse').txt('true').up();

    // Create standard service
    const standardService = service.ele('StandardService')
      .ele('Origin').txt('Standard').up()
      .ele('Destination').txt('Liverpool').up()
      .ele('UseAllStopPoints').txt('false').up();

    // Process each sheet for journey patterns
    let journeyPatternRefCounter = 1;
    let vehicleJourneyCounter = 1;

    // Add outbound pattern
    const outboundPattern = standardService
      .ele('JourneyPattern')
      .att('id', `jp_${journeyPatternRefCounter++}`)
      .ele('Direction').txt('outbound').up();

    OUTBOUND_STOPS.forEach((stop, index) => {
      outboundPattern
        .ele('JourneyPatternSection')
        .ele('JourneyPatternTimingLink')
        .ele('From')
        .ele('StopPointRef').txt(STOPS[stop].atcoCode).up()
        .ele('TimingStatus').txt('PrincipalTimingPoint').up().up()
        .ele('RunTime').txt('PT5M').up().up();
    });

    // Add return pattern
    const returnPattern = standardService
      .ele('JourneyPattern')
      .att('id', `jp_${journeyPatternRefCounter++}`)
      .ele('Direction').txt('inbound').up();

    RETURN_STOPS.forEach((stop, index) => {
      returnPattern
        .ele('JourneyPatternSection')
        .ele('JourneyPatternTimingLink')
        .ele('From')
        .ele('StopPointRef').txt(STOPS[stop].atcoCode).up()
        .ele('TimingStatus').txt('PrincipalTimingPoint').up().up()
        .ele('RunTime').txt('PT5M').up().up();
    });

    // Process timetable data from Excel
    const vehicleJourneys = doc.ele('VehicleJourneys');

    workbook.SheetNames.forEach(sheetName => {
      const sheet = workbook.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false });

      // Find the row with stop names
      const headerRowIndex = data.findIndex(row => row && row.includes('STD'));
      if (headerRowIndex === -1) return;

      // Process each column
      for (let colIndex = 0; colIndex < data[headerRowIndex].length; colIndex++) {
        if (!data[headerRowIndex][colIndex]) continue;

        const times = [];
        for (let rowIndex = headerRowIndex + 1; rowIndex < data.length; rowIndex++) {
          if (data[rowIndex] && data[rowIndex][colIndex]) {
            const formattedTime = formatTimeString(data[rowIndex][colIndex]);
            if (formattedTime) {
              times.push(formattedTime);
            }
          }
        }

        if (times.length > 0) {
          // Create vehicle journey
          const journey = vehicleJourneys
            .ele('VehicleJourney')
            .ele('PrivateCode').txt(`vj_${vehicleJourneyCounter++}`).up()
            .ele('VehicleJourneyCode').txt(`vj_${vehicleJourneyCounter}`).up()
            .ele('ServiceRef').txt(SERVICE_CODE).up()
            .ele('LineRef').txt(LINE_NAME).up()
            .ele('JourneyPatternRef').txt(`jp_${colIndex < 6 ? 1 : 2}`).up();

          // Add departure time with proper formatting
          const departureTime = times[0];
          journey.ele('DepartureTime').txt(departureTime).up();

          // Add timing links for each stop in the pattern
          const stopPattern = colIndex < 6 ? OUTBOUND_STOPS : RETURN_STOPS;
          journey.ele('VehicleJourneyTimingLink');

          // Add operating profile reference
          journey.ele('OperatingProfile')
            .ele('RegularDayType')
            .ele('DaysOfWeek')
            .ele('Christmas').up();
        }
      }
    });

    // Add StopPoints section
    const stopPoints = doc.ele('StopPoints');
    Object.entries(STOPS).forEach(([id, details]) => {
      stopPoints
        .ele('AnnotatedStopPointRef')
        .ele('StopPointRef').txt(details.atcoCode).up()
        .ele('CommonName').txt(details.commonName).up();
    });

    // Generate the XML string
    const xmlString = doc.end({ prettyPrint: true });

    // Write to file
    await fs.writeFile('ChristmasService.xml', xmlString);
    console.log('TransXChange file created successfully!');

  } catch (error) {
    console.error('Error creating TransXChange file:', error);
  }
}

// Usage
createTXC('./timetables/Xmas Timetable 2024-2025.xlsx');
