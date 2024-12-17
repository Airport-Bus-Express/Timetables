const XLSX = require('xlsx');
const fs = require('fs').promises;
const path = require('path');
const STOPS = require('./timetables/stops.json');

// GTFS Configuration
const ROUTE_ID = 'xmas_service';
const AGENCY_ID = 'agency';
const SERVICE_ID = 'xmas_2024';

// Stop definitions
// const STOPS = {
//   'STD': { stop_id: 'STD', stop_name: 'Standard', stop_lat: 0, stop_lon: 0 },
//   'STR': { stop_id: 'STR', stop_name: 'Straight', stop_lat: 0, stop_lon: 0 },
//   'BETH': { stop_id: 'BETH', stop_name: 'Bethlehem', stop_lat: 0, stop_lon: 0 },
//   'LIV': { stop_id: 'LIV', stop_name: 'Liverpool', stop_lat: 0, stop_lon: 0 }
// };

// Define the sequence of stops for outbound and return journeys
const OUTBOUND_STOPS = ['STD', 'STR', 'BETH', 'LIV', 'BETH'];
const RETURN_STOPS = ['LIV', 'BETH', 'STR', 'STD'];

async function createGTFS(inputFile) {
  try {
    // Read Excel file
    const workbook = XLSX.readFile(inputFile);

    // Initialize GTFS files content
    let stops = ['stop_id,stop_name,stop_lat,stop_lon'];
    let routes = ['route_id,agency_id,route_short_name,route_long_name,route_type'];
    let trips = ['route_id,service_id,trip_id,direction_id'];
    let stop_times = ['trip_id,arrival_time,departure_time,stop_id,stop_sequence'];
    let calendar = ['service_id,monday,tuesday,wednesday,thursday,friday,saturday,sunday,start_date,end_date'];

    // Add stops
    Object.values(STOPS).forEach(stop => {
      stops.push(`${stop.stop_id},${stop.stop_name},${stop.stop_lat},${stop.stop_lon}`);
    });

    // Add route
    routes.push(`${ROUTE_ID},${AGENCY_ID},XMAS,Christmas Special Service,3`);

    // Process each sheet (date)
    workbook.SheetNames.forEach((sheetName, sheetIndex) => {
      const sheet = workbook.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      // Find the row with stop names (usually row 4 based on the example)
      const headerRowIndex = data.findIndex(row => row && row.includes('STD'));
      if (headerRowIndex === -1) return;

      // Process each column pair (outbound and return)
      for (let colIndex = 0; colIndex < data[headerRowIndex].length; colIndex++) {
        if (!data[headerRowIndex][colIndex]) continue;

        // Process times for this column
        const times = [];
        for (let rowIndex = headerRowIndex + 1; rowIndex < data.length; rowIndex++) {
          if (data[rowIndex] && data[rowIndex][colIndex]) {
            times.push(data[rowIndex][colIndex]);
          }
        }

        if (times.length > 0) {
          // Create trip
          const tripId = `${sheetName}_${colIndex}`;
          const directionId = colIndex < 6 ? '0' : '1'; // Assuming first 6 columns are outbound
          trips.push(`${ROUTE_ID},${SERVICE_ID},${tripId},${directionId}`);

          // Add stop times
          const stopSequence = directionId === '0' ? OUTBOUND_STOPS : RETURN_STOPS;
          stopSequence.forEach((stopId, sequence) => {
            const time = times[sequence] || times[0]; // Use first time if specific stop time is missing
            if (time) {
              stop_times.push(`${tripId},${time}:00,${time}:00,${stopId},${sequence + 1}`);
            }
          });
        }
      }
    });

    // Add calendar entry (example: service runs only on specific dates)
    calendar.push(`${SERVICE_ID},0,0,0,0,0,0,0,20241224,20250101`);

    // Create output directory
    await fs.mkdir('gtfs', { recursive: true });

    // Write GTFS files
    await Promise.all([
      fs.writeFile('gtfs/stops.txt', stops.join('\n')),
      fs.writeFile('gtfs/routes.txt', routes.join('\n')),
      fs.writeFile('gtfs/trips.txt', trips.join('\n')),
      fs.writeFile('gtfs/stop_times.txt', stop_times.join('\n')),
      fs.writeFile('gtfs/calendar.txt', calendar.join('\n'))
    ]);

    console.log('GTFS files created successfully!');

  } catch (error) {
    console.error('Error creating GTFS files:', error);
  }
}

// Usage
createGTFS('./timetables/Xmas Timetable 2024-2025.xlsx');