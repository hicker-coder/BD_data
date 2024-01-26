import geopandas as gpd
import pandas as pd
from shapely.geometry import Point
from bokeh.plotting import figure
from bokeh.tile_providers import get_provider, Vendors
from bokeh.models import ColumnDataSource, Circle
from streamlit_bokeh_events import streamlit_bokeh_events
import streamlit as st
import pydeck as pdk


def find_polygons_within_buffer(lat, lon, buffer_distance_km):
    # Create a point object from the input latitude and longitude
    point = Point(lon, lat)
    point_gseries = gpd.GeoSeries([point], crs="EPSG:4326")  # Convert the Point to a GeoSeries
    point_gseries = point_gseries.to_crs("EPSG:3035")  # Reproject the GeoSeries to EPSG:3035
    point = point_gseries.iloc[0]  # Get the reprojected Point object

    geo_df = gpd.read_file('data.geojson')
    geo_df = geo_df.to_crs("EPSG:3035")
    # Create a buffer around the point
    buffer = point.buffer(buffer_distance_km * 1000)  # Multiply by 1000 to convert km to meters
    buffer_gseries = gpd.GeoSeries([buffer], crs="EPSG:3035")
    buffer = buffer_gseries.iloc[0]

    # Check which polygons intersect with the buffer
    poly_indices = geo_df[geo_df.intersects(buffer)].index

    # Get the subset of the original GeoDataFrame with the intersecting polygons
    subset = geo_df.loc[poly_indices].copy()

    # Remove polygons that are completely outside the buffer zone
    subset['intersection'] = subset.intersection(buffer)
    subset = subset[subset.area > 0].copy()

    # Calculate the percentage of each polygon's area within the buffer zone
    subset['intersection'] = subset['intersection'].to_crs("EPSG:4326")  # Reproject the intersection geometry back to EPSG:4326
    subset['area_percentage'] = subset['intersection'].area / subset['geometry'].to_crs("EPSG:4326").area

    # Read the CSV file
    df = pd.read_csv('POSAS_2022_it_Comuni.csv', sep=',', encoding='ISO-8859-1')
    # Get the comune codes from the subset GeoDataFrame
    comune_codes = subset['com_istat_code_num'].values.tolist()

    # Filter the rows in the DataFrame using the comune codes
    filtered_df = df[df['Codice comune'].isin(comune_codes)]

    # Filter the rows based on age criteria
    active_df = filtered_df[(filtered_df['Età'] >= 15) & (filtered_df['Età'] <= 64)]

    # Compute the total population and active population
    total_pop = 0
    total_active_pop = 0

    for idx, row in subset.iterrows():
        comune_code = row['com_istat_code_num']
        area_percentage = row['area_percentage']

        comune_pop = filtered_df.loc[filtered_df['Codice comune'] == comune_code]
        comune_total_pop = comune_pop[comune_pop['Età'] == 999]['Totale maschi'].sum() + comune_pop[comune_pop['Età'] == 999]['Totale femmine'].sum()
        comune_active_pop = comune_pop[(comune_pop['Età'] >= 15) & (comune_pop['Età'] <= 64)]['Totale maschi'].sum() + comune_pop[(comune_pop['Età'] >= 15) & (comune_pop['Età'] <= 64)]['Totale femmine'].sum()
        comune_active_male = comune_pop[(comune_pop['Età'] >= 15) & (comune_pop['Età'] <= 64)]['Totale maschi'].sum()
        comune_active_female = comune_pop[(comune_pop['Età'] >= 15) & (comune_pop['Età'] <= 64)]['Totale femmine'].sum()

        total_pop += comune_total_pop * area_percentage
        total_active_pop += comune_active_pop * area_percentage

        # Store total males, females, active males, and active females for this comune in the subset
        subset.loc[idx, 'Totale maschi'] = comune_pop[comune_pop['Età'] == 999]['Totale maschi'].sum()
        subset.loc[idx, 'Totale femmine'] = comune_pop[comune_pop['Età'] == 999]['Totale femmine'].sum()
        subset.loc[idx, 'Active male'] = comune_active_male
        subset.loc[idx, 'Active female'] = comune_active_female
   
    return total_pop, total_active_pop, subset[['name', 'area_percentage', 'Totale maschi', 'Totale femmine', 'Active male', 'Active female', 'Unemployment 2022', 'regional_cies', 'students',
    'lvl_primary_school_k', 'secondary_school_k','lvl_university_k', 'poverty_incidence','no_degree_wage', 'secondary_degree_wage', 'tertiary_degree_wage', 'all_levels_wage']]


# Helper function to convert latitude and longitude to Web Mercator coordinates
def create_map(lat, lon):
    view_state = pdk.ViewState(latitude=lat, longitude=lon, zoom=13, bearing=0, pitch=0)

    layer = pdk.Layer(
        "ScatterplotLayer",
        data=pd.DataFrame({"lat": [lat], "lon": [lon]}),
        get_position=["lon", "lat"],
        get_color=[255, 0, 0],
        get_radius=100,
        pickable=True
    )

    # Set the map style
    map_style = pdk.maps.MapboxStyle.DARK

    return pdk.Deck(
        map_style=map_style,
        initial_view_state=view_state,
        layers=[layer],
        tooltip={"text": "Current location: [{lon:.2f}, {lat:.2f}]", "show": True}
    )

# Streamlit app code
st.title("Geo Locator")
st.subheader("Click on the map to change the location")

if "location" not in st.session_state:
    st.session_state.location = (45.0, 0.0)

lat, lon = st.session_state.location
map = create_map(lat, lon)

# Update location when the map is clicked
new_location = streamlit_bokeh_events(
    map, events="PICK_MOVE", key="map", refresh_on_update=False, debounce_time=0)

if new_location:
    lat, lon = new_location["lat"], new_location["lon"]
    st.session_state.location = (lat, lon)
    map = create_map(lat, lon)

st.pydeck_chart(map)
st.write(f"Selected location: {st.session_state.location}")