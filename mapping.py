import streamlit as st
import pandas as pd
import folium
from geopy.geocoders import Nominatim
from geopy.distance import geodesic  # Import the geodesic function for distance calculation
from io import BytesIO
from datetime import datetime
from streamlit_folium import folium_static

# Initialize the geocoder
geolocator = Nominatim(user_agent="streamlit")

# Function to geocode based on country and postal code or city
def geocode_location(country_code, postal_code=None, city=None):
    try:
        if postal_code:
            location = geolocator.geocode(f"{postal_code}, {country_code}")
            if location:
                return location.latitude, location.longitude, postal_code, city, country_code
            
            if city:
                location = geolocator.geocode(f"{city}, {postal_code}, {country_code}")
                if location:
                    return location.latitude, location.longitude, postal_code, city, country_code
        
        if city:
            location = geolocator.geocode(f"{city}, {country_code}")
            if location:
                return location.latitude, location.longitude, None, city, country_code
        
        capital_location = geolocator.geocode(f"capital city of {country_code}")
        if capital_location:
            return capital_location.latitude, capital_location.longitude, None, None, country_code
        
        return None, None, None, None, None
    except Exception as e:
        return None, None, None, None, None

# Function to save DataFrame to Excel and return as bytes
def save_to_excel(df, original_filename):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Geocoding Data')
    output.seek(0)

    current_date = datetime.now().strftime("%m%d%Y")
    export_filename = f"{original_filename.split('.')[0]}_geocoding_details_{current_date}.xlsx"
    
    return output.getvalue(), export_filename

# Function to validate template columns
def validate_template(df, scenario):
    if scenario == "Standard visualization":
        required_columns = {'country_code', 'postal_code', 'city', 'layer'}
    elif scenario == "Supply-chain visualization":
        required_columns = {'country_code_warehouse', 'postal_code_warehouse', 'city_warehouse', 
                            'country_code_dest', 'postal_code_dest', 'city_dest', 'layer'}
    elif scenario == "Distance calculation":
        required_columns = {'country_code_orig', 'postal_code_orig', 'city_orig','country_code_dest', 'postal_code_dest', 'city_dest'}
    else:
        return False
    
    missing_columns = required_columns - set(df.columns)
    if missing_columns:
        st.error(f"Uploaded file is missing columns: {', '.join(missing_columns)}")
        return False

    return True

# Function to create the map
def create_map():
    if st.session_state.map is None:
        initial_location = [20, 0]
        st.session_state.map = folium.Map(location=initial_location, zoom_start=2)
    
    return st.session_state.map

# Streamlit app starts here
st.set_page_config(page_title="Data Visualization Tool", layout="wide")

st.title("Data Visualization Tool")

# Initialize session state variables if they don't exist
if 'layer_colors' not in st.session_state:
    st.session_state.layer_colors = {
        1: "#4D148C", 2: "#FF6200", 3: "#671CAA", 4: "#7D22C3", 5: "#932DA2", 6: "#A63685",
        7: "#B83F6A", 8: "#C74755", 9: "#D87E88", 10: "#C172AA",
    }

if 'map' not in st.session_state:
    st.session_state.map = None

if 'df' not in st.session_state:
    st.session_state.df = None

# Initialize dot size if it doesn't exist
if 'dot_size' not in st.session_state:
    st.session_state.dot_size = 2  # Set a default value for dot size

with st.sidebar:
    # Store the currently selected scenario
    selected_scenario = st.selectbox("Select a scenario", ("Standard visualization", "Supply-chain visualization", "Distance calculation"))

    # Check if the scenario has changed
    if 'scenario' in st.session_state and st.session_state.scenario != selected_scenario:
        # Clear relevant session state data when the scenario changes
        st.session_state.map = None
        st.session_state.df = None
        st.session_state.dot_size = 2  # Reset dot size if needed

    # Store the current scenario
    st.session_state.scenario = selected_scenario

    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
    create_map_button = st.button("Create", key="create")

col1, col2 = st.columns([1, 3])

if uploaded_file:
    # Read the Excel file with proper handling for leading zeros in postal codes
    dtype_mapping = {
        'postal_code': str,
        'postal_code_warehouse': str,
        'postal_code_dest': str,
        'country_code': str,
        'country_code_warehouse': str,
        'country_code_dest': str
    }

    df = pd.read_excel(uploaded_file, dtype=dtype_mapping)
    st.session_state.df = df

    if create_map_button and validate_template(df, selected_scenario):
        df['latitude'] = None
        df['longitude'] = None

        if selected_scenario == "Supply-chain visualization":
            df['warehouse_lat'] = None
            df['warehouse_lon'] = None

        total_rows = len(df)
        progress_bar_container = st.sidebar.progress(0)
        progress_text_container = st.sidebar.empty()

        location_bounds = []  # List to store all coordinates for fitting map bounds

        for index, row in df.iterrows():
            if selected_scenario == "Standard visualization":
                country_code = row['country_code']
                postal_code = row.get('postal_code')
                city = row.get('city')

                lat, lon, _, _, _ = geocode_location(country_code, postal_code, city)
                df.at[index, 'latitude'] = lat
                df.at[index, 'longitude'] = lon

                if lat is not None and lon is not None:
                    location_bounds.append([lat, lon])

            elif selected_scenario == "Supply-chain visualization":
                # Geocode destination
                country_code_dest = row['country_code_dest']
                postal_code_dest = row.get('postal_code_dest')
                city_dest = row.get('city_dest')

                lat, lon, _, _, _ = geocode_location(country_code_dest, postal_code_dest, city_dest)
                df.at[index, 'latitude'] = lat
                df.at[index, 'longitude'] = lon

                # Geocode warehouse
                country_code_warehouse = row['country_code_warehouse']
                postal_code_warehouse = row.get('postal_code_warehouse')
                city_warehouse = row.get('city_warehouse')

                warehouse_lat, warehouse_lon, _, _, _ = geocode_location(country_code_warehouse, postal_code_warehouse, city_warehouse)
                df.at[index, 'warehouse_lat'] = warehouse_lat
                df.at[index, 'warehouse_lon'] = warehouse_lon

                if warehouse_lat is not None and warehouse_lon is not None:
                    location_bounds.append([warehouse_lat, warehouse_lon])

            progress = (index + 1) / total_rows
            progress_bar_container.progress(progress)
            progress_text_container.text(f"Processing row {index + 1}/{total_rows}...")

        st.session_state.df = df

        # Create and plot on the map
        map_object = create_map()

        if selected_scenario == "Standard visualization":
            for index, row in df.iterrows():
                lat = row['latitude']
                lon = row['longitude']
                layer = row['layer']

                if lat and lon:
                    color = st.session_state.layer_colors.get(layer, "#808080")
                    folium.CircleMarker(
                        location=[lat, lon],
                        radius=st.session_state.dot_size,
                        color=color,
                        fill=True,
                        fill_color=color,
                        fill_opacity=1.0
                    ).add_to(map_object)
                    location_bounds.append([lat, lon])  # Add to bounds for zoom

            # Fit the map to the bounds of all plotted locations
            if location_bounds:
                map_object.fit_bounds(location_bounds)

            # Render the map
            folium_static(map_object)

        elif selected_scenario == "Supply-chain visualization":
            for index, row in df.iterrows():
                warehouse_lat = row['warehouse_lat']
                warehouse_lon = row['warehouse_lon']
                lat = row['latitude']
                lon = row['longitude']
                layer = row['layer']

                if warehouse_lat and warehouse_lon:
                    folium.CircleMarker(
                        location=[warehouse_lat, warehouse_lon],
                        radius=st.session_state.dot_size * 1.5,
                        color='yellow',
                        fill=True,
                        fill_color='yellow',
                        fill_opacity=1.0
                    ).add_to(map_object)
                    location_bounds.append([warehouse_lat, warehouse_lon])  # Add warehouse to bounds for zoom

                if lat and lon:
                    color = st.session_state.layer_colors.get(layer, "#808080")
                    folium.CircleMarker(
                        location=[lat, lon],
                        radius=st.session_state.dot_size,
                        color=color,
                        fill=True,
                        fill_color=color,
                        fill_opacity=1.0
                    ).add_to(map_object)
                    location_bounds.append([lat, lon])  # Add to bounds for zoom

                # Add lines connecting warehouse to destination
                if warehouse_lat and warehouse_lon and lat and lon:
                    folium.PolyLine(
                        locations=[[warehouse_lat, warehouse_lon], [lat, lon]],
                        color='grey', weight=0.5, opacity=1
                    ).add_to(map_object)

            # Fit the map to the bounds of all plotted locations
            if location_bounds:
                map_object.fit_bounds(location_bounds)

            # Render the map
            folium_static(map_object)

        # Save results and allow user to download the Excel file with distances
        if st.button("Download results"):
            result_data, result_filename = save_to_excel(df, uploaded_file.name)
            st.download_button(label="Download Excel file", data=result_data, file_name=result_filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        elif selected_scenario == "Distance calculation":
            st.warning("Test if works")
    




        # Clean the progress bars in the end
        progress_bar_container.empty()
        progress_text_container.empty()