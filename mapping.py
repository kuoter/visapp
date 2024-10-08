import streamlit as st
import pandas as pd
import folium
from geopy.geocoders import Nominatim
from geopy.distance import geodesic
from io import BytesIO
from datetime import datetime
from streamlit_folium import folium_static
import os
from zipfile import ZipFile
import streamlit.components.v1 as components

# Initialize the geocoder
geolocator = Nominatim(user_agent="streamlit")

# # Function to geocode based on country and postal code or city
# def geocode_location_old_approach(country_code, postal_code=None, city=None):
#     try:
#         if postal_code:
#             location = geolocator.geocode(f"{postal_code}, {country_code}")
#             if location:
#                 return location.latitude, location.longitude, postal_code, city, country_code
            
#             if city:
#                 location = geolocator.geocode(f"{city}, {postal_code}, {country_code}")
#                 if location:
#                     return location.latitude, location.longitude, postal_code, city, country_code
        
#         if city:
#             location = geolocator.geocode(f"{city}, {country_code}")
#             if location:
#                 return location.latitude, location.longitude, None, city, country_code
        
#         capital_location = geolocator.geocode(f"capital city of {country_code}")
#         if capital_location:
#             return capital_location.latitude, capital_location.longitude, None, None, country_code
        
#         return None, None, None, None, None
#     except Exception as e:
#         return None, None, None, None, None

# Function to geocode based on country code, postal code, and city with fallback to nearest location
def geocode_location(country_code, postal_code=None, city=None):
    try:
        # Helper function to check if geocoded location is in the correct country
        def is_in_correct_country(location, country_code):
            if location and hasattr(location, 'raw'):
                # Ensure the country code matches
                address_details = location.raw.get('address', {})
                return address_details.get('country_code', '').upper() == country_code.upper()
            return False
        
        # Helper function to attempt nearest location reverse geocoding
        def find_nearest_location(location):
            if location:
                # Attempt reverse geocoding near the found location
                reverse_location = geolocator.reverse((location.latitude, location.longitude), exactly_one=True)
                if reverse_location and is_in_correct_country(reverse_location, country_code):
                    return reverse_location.latitude, reverse_location.longitude, postal_code, city, country_code
            return None, None, None, None, None

        # Try geocoding by city, postal code, and country (most specific)
        if city and postal_code:
            location = geolocator.geocode(f"{city}, {postal_code}, {country_code}")
            if location and is_in_correct_country(location, country_code):
                return location.latitude, location.longitude, postal_code, city, country_code
            # Attempt nearest location if unsuccessful
            nearest = find_nearest_location(location)
            if nearest:
                return nearest
        
        # Try geocoding by postal code and country
        if postal_code:
            location = geolocator.geocode(f"{postal_code}, {country_code}")
            if location and is_in_correct_country(location, country_code):
                return location.latitude, location.longitude, postal_code, None, country_code
            # Attempt nearest location if unsuccessful
            nearest = find_nearest_location(location)
            if nearest:
                return nearest
        
        # Try geocoding by city and country
        if city:
            location = geolocator.geocode(f"{city}, {country_code}")
            if location and is_in_correct_country(location, country_code):
                return location.latitude, location.longitude, None, city, country_code
            # Attempt nearest location if unsuccessful
            nearest = find_nearest_location(location)
            if nearest:
                return nearest
        
        # If all else fails, geocode the capital city of the country
        capital_location = geolocator.geocode(f"capital city of {country_code}")
        if capital_location and is_in_correct_country(capital_location, country_code):
            return capital_location.latitude, capital_location.longitude, None, None, country_code
        
        # Return None if no valid geocode results were found
        return None, None, None, None, None

    except Exception as e:
        # Handle exceptions gracefully
        print(f"Geocoding error: {e}")
        return None, None, None, None, None

def save_to_excel(df, original_filename):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Geocoding Data')
    output.seek(0)  # Important to reset the pointer to the start of the BytesIO buffer

    current_date = datetime.now().strftime("%m%d%Y")
    export_filename = f"{original_filename.split('.')[0]}_geocoding_details_{current_date}.xlsx"

    return output, export_filename  # Return the BytesIO object itself, not getvalue()

# Function to validate template columns
def validate_template(df, scenario):
    if scenario == "Standard visualization":
        required_columns = {'country_code', 'postal_code', 'city', 'layer'}
    elif scenario == "Supply-chain visualization":
        required_columns = {'country_code_warehouse', 'postal_code_warehouse', 'city_warehouse', 
                            'country_code_dest', 'postal_code_dest', 'city_dest', 'layer'}
    elif scenario == "Distance calculation":
        required_columns = {'country_code_orig', 'postal_code_orig', 'city_orig','country_code_dest', 'postal_code_dest', 'city_dest'}
    elif scenario == "Volume visualization":
        required_columns = {'country_code', 'postal_code', 'city', 'volume'}
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

def zip_templates_folder():
    # Create a zip file in memory
    zip_buffer = BytesIO()
    with ZipFile(zip_buffer, 'w') as zip_file:
        templates_folder = os.path.join(os.path.dirname(__file__), 'templates')  # Path to the templates folder
        for filename in os.listdir(templates_folder):
            file_path = os.path.join(templates_folder, filename)
            if os.path.isfile(file_path):
                zip_file.write(file_path, os.path.basename(file_path))  # Add file to zip
    zip_buffer.seek(0)  # Move to the start of the BytesIO buffer
    return zip_buffer.getvalue()

def scale_dot_size(volume, min_volume, max_volume):
    """Dynamically scale the dot size based on the volume value."""
    min_size = 2  # Minimum dot size
    max_size = 15  # Maximum dot size
    if max_volume == min_volume:  # Prevent division by zero
        return min_size
    return min_size + (max_size - min_size) * ((volume - min_volume) / (max_volume - min_volume))

# Streamlit app starts here
st.set_page_config(page_title="Data Visualization Tool", layout="wide")

st.title("Data Visualization Tool")

progress_bar_container = st.empty() 
progress_text_container = st.empty()

### Version ###
st.sidebar.markdown("---")
st.sidebar.markdown("<p style='text-align: center; font-size:10px; margin-top:-20px;'>Version: 1.3.0</p>", unsafe_allow_html=True)
### Version ###

# Initialize session state variables if they don't exist
if 'layer_colors' not in st.session_state:
    st.session_state.layer_colors = {
        1: "#4D148C", 2: "#FF6200", 3: "#671CAA", 4: "#7D22C3", 5: "#932DA2", 6: "#A63685",
        7: "#B83F6A", 8: "#C74755", 9: "#D87E88", 10: "#C172AA",
        # Add more colors here #
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
    selected_scenario = st.selectbox("Select a scenario", ("Standard visualization","Volume visualization", "Supply-chain visualization", "Distance calculation"))

    # Check if the scenario has changed
    if 'scenario' in st.session_state and st.session_state.scenario != selected_scenario:
        # Clear relevant session state data when the scenario changes
        st.session_state.map = None
        st.session_state.df = None
        st.session_state.dot_size = 2  # Reset dot size if needed

    # Store the current scenario
    st.session_state.scenario = selected_scenario

    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
    create_map_button = st.button("CREATE", key="create")

col1, col2 = st.columns([1, 3])

st.sidebar.markdown("<br>" * 3, unsafe_allow_html=True) # Make some space
st.sidebar.download_button(
    label="Download templates",
    data=zip_templates_folder(),
    file_name="templates.zip",
    mime="application/zip"
)

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

            elif selected_scenario == "Volume visualization":
                # Find min and max volume for smart scaling
                min_volume = df['volume'].min()
                max_volume = df['volume'].max()

                country_code = row['country_code']
                postal_code = row.get('postal_code')
                city = row.get('city')

                lat, lon, _, _, _ = geocode_location(country_code, postal_code, city)
                df.at[index, 'latitude'] = lat
                df.at[index, 'longitude'] = lon

                if lat is not None and lon is not None:
                    location_bounds.append([lat, lon])

            elif selected_scenario == "Supply-chain visualization":
                country_code_dest = row['country_code_dest']
                postal_code_dest = row.get('postal_code_dest')
                city_dest = row.get('city_dest')

                lat, lon, _, _, _ = geocode_location(country_code_dest, postal_code_dest, city_dest)
                df.at[index, 'latitude'] = lat
                df.at[index, 'longitude'] = lon

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

        st.session_state.df = df

        # Create and plot on the map
        map_object = create_map()

        # Inside your "Standard visualization" scenario
        if selected_scenario == "Standard visualization":
            plotted_layers = set()  # Track plotted layers

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

                    # Add the layer to the set of plotted layers
                    plotted_layers.add(layer)

            # Fit the map to the bounds of all plotted locations
            if location_bounds:
                map_object.fit_bounds(location_bounds)

            # Render the map on the left side
            col1, col2 = st.columns([2, 1])
            with col1:
                folium_static(map_object)

            # Define the legend HTML outside the map and display it on the right
            with col2:
                legend_html = """
                    <div style='border:1px solid grey; padding: 1px; border-radius: 1px; background-color:black; width: 150px;font-family: "Inter", sans-serif; color: white; font-size: 13px;'>
                    <h4>LEGEND</h4>
                """

                # Append each plotted layer and its color to the HTML string
                for layer in plotted_layers:
                    color = st.session_state.layer_colors.get(layer, "#808080")
                    legend_html += f"<div style='margin-bottom: 1px;'><span style='background-color:{color}; width: 10px; height: 10px; border-radius: 50%; display: inline-block; margin-right: 1px;'></span>  {layer} day(s)</div>"

                legend_html += "</div>"

                # Render the legend HTML in the right column
                components.html(legend_html, height=150 + len(plotted_layers) * 20)  # Adjust height dynamically

            progress_bar_container.empty()

        if selected_scenario == "Volume visualization":
                    plotted_layers = set()  # Track plotted layers

                    for index, row in df.iterrows():
                        lat = row['latitude']
                        lon = row['longitude']
                        volume = row['volume']

                        if lat and lon:
                            color = st.session_state.layer_colors.get(volume, "#808080")
                            size = scale_dot_size(volume, min_volume, max_volume)
                            folium.CircleMarker(
                                location=[lat, lon],
                                radius=size,
                                color=color,
                                fill=True,
                                fill_color="808080",
                                fill_opacity=1.0
                            ).add_to(map_object)
                            location_bounds.append([lat, lon])  # Add to bounds for zoom

                            # Add the layer to the set of plotted layers
                            plotted_layers.add(volume)

                    # Fit the map to the bounds of all plotted locations
                    if location_bounds:
                        map_object.fit_bounds(location_bounds)

                    # Render the map on the left side
                    col1, col2 = st.columns([2, 1])
                    with col1:
                        folium_static(map_object)

                    progress_bar_container.empty()

        elif selected_scenario == "Supply-chain visualization":
            plotted_layers = set()  # Track plotted layers
            location_bounds = []  # Initialize location bounds

            for index, row in df.iterrows():
                warehouse_lat = row['warehouse_lat']
                warehouse_lon = row['warehouse_lon']
                lat = row['latitude']
                lon = row['longitude']
                layer = row['layer']

                # Plot the warehouse marker
                if warehouse_lat and warehouse_lon:
                    folium.CircleMarker(
                        location=[warehouse_lat, warehouse_lon],
                        radius=st.session_state.dot_size * 1.5,
                        color='yellow',
                        fill=True,
                        fill_color='yellow',
                        fill_opacity=1.0
                    ).add_to(map_object)
                    location_bounds.append([warehouse_lat, warehouse_lon])

                # Plot the regular location markers
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

                    # Add the layer to the set of plotted layers
                    plotted_layers.add(layer)

            # Fit the map to the bounds of all plotted locations
            if location_bounds:
                map_object.fit_bounds(location_bounds)

            # Render the map on the left side
            col1, col2 = st.columns([2, 1])
            with col1:
                folium_static(map_object)

            # Define the legend HTML outside the map and display it on the right
            with col2:
                legend_html = """
                    <div style='border:1px solid grey; padding: 1px; border-radius: 1px; background-color:black; width: 150px; font-family: "Inter", sans-serif; color: white; font-size: 13px;'>
                    <h4>LEGEND</h4>
                    <div style='margin-bottom: 1px;'>
                        <span style='background-color:yellow; width: 10px; height: 10px; border-radius: 50%; display: inline-block; margin-right: 1px;'></span>
                        Shipping location
                    </div>
                """

                # Append each plotted layer and its color to the HTML string
                for layer in plotted_layers:
                    color = st.session_state.layer_colors.get(layer, "#808080")
                    legend_html += f"<div style='margin-bottom: 1px;'><span style='background-color:{color}; width: 10px; height: 10px; border-radius: 50%; display: inline-block; margin-right: 1px;'></span>  {layer} day(s)</div>"

                legend_html += "</div>"

                # Render the legend HTML in the right column
                components.html(legend_html, height=150 + len(plotted_layers) * 20)  # Adjust height dynamically

            progress_bar_container.empty()

        elif selected_scenario == "Distance calculation":

            if 'latitude' not in df.columns or 'longitude' not in df.columns:
                df['orig_latitude'] = None
                df['orig_longitude'] = None
                df['dest_latitude'] = None
                df['dest_longitude'] = None

        df['distance_km'] = None

        total_rows = len(df)

        for index, row in df.iterrows():

            # Geocode the origin
            country_code_orig = row['country_code_orig']
            postal_code_orig = row.get('postal_code_orig')
            city_orig = row.get('city_orig')

            orig_lat, orig_lon, _, _, _ = geocode_location(country_code_orig, postal_code_orig, city_orig)
            df.at[index, 'orig_latitude'] = orig_lat
            df.at[index, 'orig_longitude'] = orig_lon

            # Geocode the destination
            country_code_dest = row['country_code_dest']
            postal_code_dest = row.get('postal_code_dest')
            city_dest = row.get('city_dest')

            dest_lat, dest_lon, _, _, _ = geocode_location(country_code_dest, postal_code_dest, city_dest)
            df.at[index, 'dest_latitude'] = dest_lat
            df.at[index, 'dest_longitude'] = dest_lon

            # Calculate the distance if both origin and destination are available
            if orig_lat and orig_lon and dest_lat and dest_lon:
                distance = geodesic((orig_lat, orig_lon), (dest_lat, dest_lon)).kilometers
                df.at[index, 'distance_km'] = int(distance)  # Convert to integer to remove decimal points

            # Update the progress bar
            progress_bar_container.progress(progress)

        # Display the results in Streamlit
        st.dataframe(df[['country_code_orig', 'postal_code_orig', 'city_orig',
                     'country_code_dest', 'postal_code_dest', 'city_dest',
                     'distance_km']])

        # Enable users to download the results as an Excel file
        result_data, result_filename = save_to_excel(df, uploaded_file.name)
        st.download_button(label="Download raw data", data=result_data, file_name=result_filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # Clean the progress bars in the end
        progress_bar_container.empty()
