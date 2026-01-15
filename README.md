# Distance Calculator for Brazilian Cities Routes

This Python application calculates distances between origins and destinations for routes in Brazilian cities using the OpenRouteService API. It processes an Excel file containing route data and outputs the distances in kilometers.

## Features

- Reads route data from an Excel file (`routes.xlsx`)
- Calculates distances using OpenRouteService API
- Handles multiple origins dynamically
- Processes data in chunks to respect API limits
- Provides a simple GUI with progress tracking
- Outputs results to a new Excel file (`routes_with_distance.xlsx`)

## Prerequisites

- Python 3.x
- Required Python packages: `pandas`, `requests`, `tkinter` (tkinter is included with Python)
- OpenRouteService API key

## Installation

1. Clone or download this repository.
2. Install the required packages:
   ```
   pip install pandas requests
   ```
3. Obtain an API key from [OpenRouteService](https://openrouteservice.org/).
4. Create a `credencial.json` file in the project directory with the following structure:
   ```json
   {
     "api_key": "your_api_key_here",
     "url": "https://api.openrouteservice.org/v2/matrix/driving-car"
   }
   ```

## Usage

1. Prepare your input Excel file (`routes.xlsx`) with the following columns:
   - `Origin`: Name of the origin city
   - `Long|Lat`: Coordinates in the format "longitude|latitude" (e.g., "-46.6333|-23.5505")
   - `Longitude`: Destination longitude
   - `Latitude`: Destination latitude

2. Run the application:
   ```
   python main.py
   ```

3. Click the "Start Processing" button in the GUI.

4. The application will process the data and display progress. Once complete, it will save the results to `routes_with_distance.xlsx` with an additional `distance_km` column.

## Testing

A simple test script `test.py` is provided to verify the API connection. Run it with:
```
python test.py
```

## Configuration

- `CHUNK_SIZE`: Number of destinations processed per API call (default: 50)
- `MAX_RETRIES`: Maximum number of API retry attempts (default: 5)
- `BACKOFF_FACTOR`: Exponential backoff factor for retries (default: 2)

## Notes

- The application respects the OpenRouteService free tier limits by adding a 1-second delay between requests.
- Ensure your API key has sufficient credits for the number of requests.
- The GUI uses Tkinter and runs in a separate thread to keep the interface responsive.

## License

This project is open-source. Please check the license file if available.
