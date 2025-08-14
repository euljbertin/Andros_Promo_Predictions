
import streamlit as st
from snowflake.snowpark.context import get_active_session
from io import BytesIO
from datetime import datetime
import pandas as pd
import openpyxl


# Write directly to the app
st.set_page_config(layout='wide')
st.title("🧮 Promotions Predictions App")

# Get the current credentials
SESSION = get_active_session()

# Constantes de configurations 
CONFIG = {
    "DATABASE": "CH_PRD_DB",
    "SCHEMA": "DWH_WRK",
    "STAGE": "CORTEX_ANALYST_PROJECT",
    "PROMO_TABLE": "CH_PRD_DB.DWH_WRK.BASELINEPREDICTION_REF_PROMOTION",
    "CATEGORY_TABLE": "CH_PRD_DB.DWH_WRK.MAPPING_SEGMENT_CATEGORY",
    "PROMO_FORECAST_TABLE": "CH_PRD_DB.DWH_WRK.BASELINEPREDICTION_REF_PROMOTION_FORECAST",
    "PROMO_FORECAST_KPI_TABLE": "CH_PRD_DB.DWH_WRK.BASELINEPREDICTION_REF_PROMOTION_FORECAST_KPI"
}

def style():
    """Gestion des styles du front de l'app streamlit."""
    style = """
        <style>
            @import url('http://www.mostardesign.com');
            @import url('https://fonts.cdnfonts.com/css/poppins');
            
            .stMain{
                background-color: rgba(255, 255, 255, 0.5);
                color : #143F49 !important;
                font-family: FilsonPro, sans-serif;
                margin: 0;
                background-image: url('https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcThL15R22HA6nwDsNdHI7q88gY3-fEyQQrfIiywtN4jQTVd63fs');
                background-size: cover;
                background-position: center;
                background-attachment: fixed; 
            }
            .stMain::before {
                content: "";
                position: absolute;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background-color: rgba(255, 255, 255, 0.98); 
            }
            
            /* Change background when screen width is less than 768px (tablets & phones) */
            @media (max-width: 768px) {
                .stMain {
                    background-color: white; /* Darker background */
                    background-image: none; /* Remove the background image */
                }
            
                .stMain::before {
                    background-color: white; /* Adjust overlay color */
                }
            }
            
            /* Change background for very small screens (mobile devices) */
            @media (max-width: 480px) {
                .stMain {
                    background-color: white; /* Even darker */
                }
            
                .stMain::before {
                    background-color: white;
                }
            }
            
            h1 {
                color : #E20E17 !important;
                font-family: "filson-pro", sans-serif;
                font-weight: 600;
                font-style: normal;
                font-size: 40px;
            }
            p {
                color : black !important;
                font-family: "filson-pro", sans-serif;
                font-weight: 400;
                font-style: normal;
            }
            [data-testid=stBottomBlockContainer] {
                background-color: white !important;
                color : #143F49 !important;
                font-family: FilsonPro, sans-serif;
            }
            [data-testid=stBottom] {
                background-color: white !important;
                color : #143F49 !important;
                font-family: FilsonPro, sans-serif;
            }
            [data-testid=stBottom] > :nth-child(1) {
              background-color: transparent !important;
            }
            [data-testid=stChatInput] {
                background-color: #F8C3C5 !important;
                border: transparent !important;
            }
            [data-testid=stChatMessage][data-role="assistant"] {
                background-color: rgba(226, 14, 23, 0.15) !important;
                padding: 1rem 3em 1rem 1rem;
            }
            [data-testid=stChatMessage]:first-of-type {
                background-color: #FFE5E6 !important;
                color: white !important;
                border-radius: 8px !important;
                padding: 1rem 3em 1rem 1rem;
            }
    
            /* Target assistant messages */
            [data-testid=stChatMessage]:nth-of-type(even) {
                background-color: #FFE5E6 !important;
                color: black !important;
                padding: 1rem 3em 1rem 1rem;
            }
            /* Target the container that contains the specific header */
            div:has(> #feedback-header) div[data-testid="stVerticalBlockBorderWrapper"] {
                border: 2px solid red !important;  /* Red border */
                border-radius: 8px !important;  /* Optional: Rounded edges */
                padding: 10px !important;  /* Optional: Adjust padding */
            }
            [data-testid=stExpander] details{
                border-color: #F8C3C5;
            }
            [data-testid=stRadio] p{
                color: black !important;
                font-weight: bold;
                font-size: 15px;
            }
            [data-testid=stRadio] [role=radiogroup]{
                padding-left: 140px;
            }
            [data-testid=stButton] button{
                background-color: white;
                border: transparent;
            }
            [data-testid=stButton] p{
                color: black !important;
                font-weight: 600;
                font-size: 13px;
            }
            
            [data-testid=stButton] {
                color: #E20E17 !important;
                font-weight: bold;
            }
            [data-testid="stBaseButton-primary"] {
                background-color: #E20E17 !important;
                border: none !important;
            }
            [data-testid="stBaseButton-primary"] p {
                color: white !important;
            }
            
            [data-baseweb="tag"] {
                background-color: #E20E17 !important;
            }

            [data-baseweb="tag"] path{
                color: white !important;
            }
            
            path{
                color: #E20E17 !important;
            }

            
            /* SIDEBAR */
            [data-testid=stSidebar] {
                background-color: #E20E17;
                color: white !important;
            }
            [data-testid=stSidebar] h4 {
                color: white;
            }
            [data-testid=stSidebar] h5 {
                color: #FFE5E6;
                font-weight: lighter;
                font-size: smaller;
            }
            [data-testid=stSidebar] p {
                color: white !important;
            }
            [data-testid=stSidebar] a {
                color: white !important;
                font-size: 14px;
            }
            [data-testid=stSidebar] details {
                background-color: rgba(255, 229, 230, 0.3) !important;
                color: white;
            }
            
            [data-testid=stSidebar] details p{
                color: white !important;
                font-weight: normal;
                font-size: 14px ;
            }
            
            [data-testid=stSidebar] details h6{
                color: white !important;
                font-weight: light !important;
                font-size: smaller;
            }
            [data-testid=stSidebar] li {
                font-size: smaller;
            }
            [data-testid=stSidebar] p {
                font-size: smaller;
            }
            [data-testid=stSidebar] .stButton{
                color: #E20E17 !important;
                font-weight: normal !important;
            }
            [data-testid=stSidebar] .stButton p{
                color: #E20E17 !important;
                font-weight: bold !important;
                font-size: 14px;
            }
            
            [data-testid=stSidebar] [data-testid=stExpander] summary p{
                color: white !important;
                font-weight: bold !important;
                font-size: 14px;
            }
            
            [data-testid=stExpanderDetails] details p{
                color: black !important;
                font-weight: lighter !important;
                font-size: 14px;
            }
            
            /* Target the sidebar collapse button */
            [data-testid="stSidebarCollapseButton"] button {
                color: red !important;  /* Change icon color */
                background-color: white !important;  /* Change background color */
                border-radius: 5px; /* Optional: Round edges */
            }
    
            /* Optional: Change color on hover */
            [data-testid="stSidebarCollapseButton"] button:hover {
                background-color: #FFE5E6 !important;
                color: white !important;
            }

            [data-testid="stPopoverBody"] button {
                font-size: 1.8em !important; /* Ajustez cette valeur pour la taille de l'émoji */
                background-color: transparent !important; /* Rend le fond transparent */
                border: none !important; /* Supprime la bordure du bouton */
                color: #E20E17 !important; /* Couleur de l'émoji/texte du bouton */
                line-height: 1 !important; /* Aide à l'alignement vertical de l'émoji */
                padding: 0px !important; /* Réduit le padding autour de l'émoji */
                margin: 0px !important; 
            }
            [data-testid="stPopoverBody"] button:hover {
                background-color: rgba(226, 14, 23, 0.1) !important; /* Effet de survol subtil */
            }
            /* Ajustement pour l'alignement du popover dans la colonne */
            div[data-testid="stPopoverBody"] > div[data-testid="stPopover"] {
                display: flex; /* Utilise flexbox pour centrer */
                justify-content: flex-start; /* Centre horizontalement */
                align-items: flex-start !important;
                height: 100%; /* Occupe toute la hauteur de la colonne */
            }
                

            
            /* DARK MODE */
            @media (prefers-color-scheme: dark) {
                .stMain{
                    background-color: rgba(255, 255, 255, 0.5);
                    color : #143F49 !important;
                    font-family: FilsonPro, sans-serif;
                    margin: 0;
                    background-image: url('https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcThL15R22HA6nwDsNdHI7q88gY3-fEyQQrfIiywtN4jQTVd63fs');
                    background-size: cover;
                    background-position: center;
                    background-attachment: fixed; 
                }
                .stMain::before {
                    content: "";
                    position: absolute;
                    top: 0;
                    left: 0;
                    width: 100%;
                    height: 100%;
                    background-color: rgba(0, 0, 0, 0.98); 
                }
                
                /* Change background when screen width is less than 768px (tablets & phones) */
                @media (max-width: 768px) {
                    .stMain {
                        background-color: black; /* Darker background */
                        background-image: none; /* Remove the background image */
                    }
                
                    .stMain::before {
                        background-color: black; /* Adjust overlay color */
                    }
                }
                
                h1 {
                    color : #E20E17 !important;
                    font-family: "filson-pro", sans-serif;
                    font-weight: 600;
                    font-style: normal;
                    font-size: 25px;
                }
                p {
                    color : white !important;
                    font-family: "filson-pro", sans-serif;
                    font-weight: 400;
                    font-style: normal;
                }
                [data-testid=stBottomBlockContainer] {
                    background-color: black !important;
                    color : black !important;
                    font-family: FilsonPro, sans-serif;
                }
                [data-testid=stBottom] {
                    background-color: black !important;
                    color : black !important;
                    font-family: FilsonPro, sans-serif;
                }
                [data-testid=stBottom] > :nth-child(1) {
                  background-color: transparent !important;
                }
                [data-testid=stChatInput] {
                    background-color: #F8C3C5 !important;
                    border: transparent !important;
                }
                textarea[aria-label="Wie lautet deine Frage?"]:not(:placeholder-shown):focus{
                 color: black;
                }
                [data-testid=stChatMessage][data-role="assistant"] {
                    background-color: rgba(226, 14, 23, 0.15) !important;
                    padding: 1rem 3em 1rem 1rem;
                }
                [data-testid=stChatMessage]:first-of-type {
                    background-color: #143F49 !important;
                    color: white !important;
                    border-radius: 8px !important;
                    padding: 1rem 3em 1rem 1rem;
                }
        
                /* Target user messages */
                [data-testid=stChatMessage]:nth-of-type(even) {
                    background-color: #143F49 !important;
                    color: black !important;
                    padding: 1rem 3em 1rem 1rem;
                }
                /* Target the container that contains the specific header */
                div:has(> #feedback-header) div[data-testid="stVerticalBlockBorderWrapper"] {
                    border: 2px solid red !important;  /* Red border */
                    border-radius: 8px !important;  /* Optional: Rounded edges */
                    padding: 10px !important;  /* Optional: Adjust padding */
                }
                [data-testid=stExpander] details{
                    border-color: #F8C3C5;
                }
                [data-testid=stRadio] p{
                    color: #E20E17 !important;
                    font-weight: bold;
                    font-size: 15px;
                }
                [data-testid=stRadio] [role=radiogroup]{
                    padding-left: 140px;
                }
                [data-testid=stButton] button{
                    background-color: white;
                    border: transparent;
                }
                [data-testid=stButton] p{
                    color: black !important;
                    font-weight: 600;
                    font-size: 13px;
                }
                
                [data-testid=stButton] {
                    color: #E20E17 !important;
                    font-weight: bold;
                }
                [data-testid="stBaseButton-primary"] {
                    background-color: #E20E17 !important;
                    border: none !important;
                }
                [data-testid="stBaseButton-primary"] p {
                    color: white !important;
                }
                
                path{
                    color: #E20E17 !important;
                }
    
                
                /* SIDEBAR */
                [data-testid=stSidebar] {
                    background-color: #E20E17;
                    color: white !important;
                }
                [data-testid=stSidebar] h4 {
                    color: white;
                }
                [data-testid=stSidebar] h5 {
                    color: #FFE5E6;
                    font-weight: lighter;
                    font-size: smaller;
                }
                [data-testid=stSidebar] p {
                    color: white !important;
                }
                [data-testid=stSidebar] a {
                    color: white !important;
                    font-size: 14px;
                }
                [data-testid=stSidebar] details {
                    background-color: rgba(255, 229, 230, 0.3) !important;
                    color: white;
                }
                
                [data-testid=stSidebar] details p{
                    color: white !important;
                    font-weight: normal;
                    font-size: 14px ;
                }
                
                [data-testid=stSidebar] details h6{
                    color: white !important;
                    font-weight: light !important;
                    font-size: smaller;
                }
                [data-testid=stSidebar] li {
                    font-size: smaller;
                }
                [data-testid=stSidebar] p {
                    font-size: smaller;
                }
                [data-testid=stSidebar] .stButton{
                    color: #E20E17 !important;
                    font-weight: normal !important;
                }
                [data-testid=stSidebar] .stButton p{
                    color: #E20E17 !important;
                    font-weight: bold !important;
                    font-size: 14px;
                }
                
                [data-testid=stSidebar] [data-testid=stExpander] summary p{
                    color: white !important;
                    font-weight: bold !important;
                    font-size: 14px;
                }
                
                [data-testid=stExpanderDetails] details p{
                    color: black !important;
                    font-weight: lighter !important;
                    font-size: 14px;
                }
                
                /* Target the sidebar collapse button */
                [data-testid="stSidebarCollapseButton"] button {
                    color: red !important;  /* Change icon color */
                    background-color: white !important;  /* Change background color */
                    border-radius: 5px; /* Optional: Round edges */
                }
        
                /* Optional: Change color on hover */
                [data-testid="stSidebarCollapseButton"] button:hover {
                    background-color: #FFE5E6 !important;
                    color: white !important;
            }
            }
        </style>
    
    """
    st.markdown(style, unsafe_allow_html=True)
    _, logo = st.columns((7,1))
    with logo :
        st.image("https://andros-asia.com/wp-content/uploads/andros-logo.png", width=100)
    return None
    
def delete_rows_table(table, ids_to_delete):
    """
    Delete rows from Snowflake table based on a list of rows deleted by the user

    Args:
        rows_to_delete_df (str): Name of the table to update.
    """
    # Get the raw connection object from the session.
    conn = SESSION.connection
        
    # Create a cursor object for executing SQL queries.
    cursor = conn.cursor()
    
    for id in ids_to_delete:
        
        # Assuming the table has an 'ID' column as the unique identifier
        delete_query = f"""
        DELETE FROM {table}
        WHERE FORECAST_PROMOTION_ID = {id};
        """
        cursor.execute(delete_query)

    # Commit changes
    cursor.execute("COMMIT")

    # Close the cursor
    cursor.close()

    
def update_snowflake_table(table, edited_df, numeric_columns=['']):
    """
    Updates a Snowflake table based on a given DataFrame.

    Args:
        table (str): Name of the table to update.
        edited_df (pd.DataFrame): DataFrame containing the new values.
        numeric_columns (list, optional): List of columns that should not be wrapped in quotes.
    """
    # Get the raw connection object from the session.
    conn = SESSION.connection
        
    # Create a cursor object for executing SQL queries.
    cursor = conn.cursor()

    columns = edited_df.columns

    for _, row in edited_df.iterrows():
        set_clauses = []
        
        for col in columns:
            value = row[col]
            
            # Format value based on data type
            if pd.isna(value):  # Handle NaN values
                formatted_value = "NULL"
            elif col in numeric_columns:  # Numeric columns (no quotes)
                formatted_value = str(value)
            else:  # String columns (wrapped in quotes)
                formatted_value = f"'{value}'"
            
            set_clauses.append(f"{col} = {formatted_value}")

        set_query = ", ".join(set_clauses)
        
        # Assuming the table has an 'ID' column as the unique identifier
        update_query = f"""
        UPDATE {table}
        SET {set_query}
        WHERE PROMOTION_ID = '{row['PROMOTION_ID']}';
        """
        cursor.execute(update_query)

    # Commit changes
    cursor.execute("COMMIT")

    # Close the cursor
    cursor.close()

def update_table_no_truncate(table, edited_df, numeric_columns=None):
    """
    Updates a Snowflake table with data from a pandas DataFrame without truncating the table.

    Args:
        - table (str): Snowflake table name.
        - edited_df (pandas.DataFrame): DataFrame containing rows to insert.
        - numeric_columns (list, optional): List of numeric column names to be inserted without quotes.
    """
    if numeric_columns is None:
        numeric_columns = []

    # Get the raw connection object from the session.
    conn = SESSION.connection
    cursor = conn.cursor()

    # Get column names
    columns = edited_df.columns.tolist()
    col_list = ', '.join(columns)  # Format as "col1, col2, col3"

    # Use `?` as the placeholder for Snowflake
    placeholders = ', '.join(['?'] * len(columns))
    insert_query = f"INSERT INTO {table} ({col_list}) VALUES ({placeholders})"

    # Convert DataFrame rows to a list of tuples
    values = [
        tuple(None if pd.isna(row[col]) else row[col] for col in columns)
        for _, row in edited_df.iterrows()
    ]
    try:
        # Execute batch insert
        cursor.executemany(insert_query, values)

        # Commit the transaction
        cursor.execute("COMMIT")
    
    except Exception as e:
        st.write(f"Error executing SQL: {e}")
    
    finally:
        # Close cursor
        cursor.close()

def save_as_excel_button(df, df_name):
        """
        Generate a Streamlit button to download a DataFrame as an Excel file.

        Args:
        - df (DataFrame): DataFrame to be saved.
        - df_name (str): Name of the file (without extension).
        """

        # Round all numbers so they only have two numbers after the comma
        df = df.round(2)

        @st.cache_data # Cache the function to avoid redundant processing.
        def to_excel(df):
            """
            Convert the DataFrame to an Excel file using an in-memory buffer.

            Args:
            - df (DataFrame): DataFrame to convert.

            Returns:
            - Bytes: Excel data in memory.
            """
            # Create an in-memory buffer for storing the Excel data.
            output = BytesIO()
            # Use openpyxl as the engine for writing Excel files.
            writer = pd.ExcelWriter(output, engine='openpyxl')
            df.to_excel(writer, index=False, sheet_name=df_name) # Write data to the first sheet.
            writer.close() # Close the writer to finalize the data.
            # Retrieve the processed Excel data.
            processed_data = output.getvalue()
            return processed_data 
        
        # Generate the Excel data from the DataFrame.
        excel_data = to_excel(df)
        
        # Create a download button in Streamlit for downloading the Excel file.
        _, save_as_excel_col = st.columns([3,1])
        with save_as_excel_col:
            st.download_button(
                label="Download data as Excel file",
                data=excel_data,
                file_name=f"{df_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

def body():
    """
    Main method to define the user interface and behavior of the Promo Calendar Input App.
    """
    style()
    st.write("")

    st.write("")
    user_action = st.radio(
            "What do you wish to do ?",
            ["Review predicted promotions", "Predict a new promotion"],
            horizontal=True,
            index=None,
    )

    #Initialize session_state
    if 'was_saved' not in st.session_state:
        st.session_state.was_saved = False
    else:
        st.session_state.was_saved = False
        
    if user_action:
        if user_action !=  "Review predicted promotions":
            st.subheader("Create your promotion")
            
        # Path to the logs table used for tracking changes
        promo_table = CONFIG['PROMO_TABLE']
        categories_table = CONFIG['CATEGORY_TABLE']
        predicted_promo_table = CONFIG['PROMO_FORECAST_KPI_TABLE']
        
        # Load the content of the tabkes into dataframes
        loaded_table = SESSION.table(promo_table)
        promo_df = loaded_table.to_pandas()
        
        categories_table = SESSION.table(categories_table)
        categories_df = categories_table.to_pandas()
        
        loaded_table = SESSION.table(predicted_promo_table)
        predicted_promo_df = loaded_table.to_pandas()
            
        current_year = datetime.now().year
        years = list(range(2023, current_year+3))
        if user_action ==  "Review predicted promotions":
            years += ['All']
            
        weeks = list(range(1, 52+1))
        
        promo_types = sorted(list(set(promo_df["PROMOTION_TYPE"])))
        if user_action ==  "Review predicted promotions":
            promo_types.insert(0, 'Any')
        else: 
            promo_types.insert(0, 'Other')
            
        duration_types = ['Tag(e)', 'Woche(n)']
        if user_action ==  "Review predicted promotions":
            duration_types = ['Any', 'Tag(e)', 'Woche(n)']
            
        regions = set(promo_df['REGION'])
        regions = list(set(part for region in regions for part in region.split('+')))
        regions = sorted(regions, key=str.casefold)
        
        mechanisms = sorted(list(set(promo_df['MECHANISM'])))
        if user_action ==  "Review predicted promotions":
            mechanisms.insert(0, 'Any')
        else: 
            mechanisms.insert(0, 'Other')
            
        brands = sorted(list(set(promo_df["BRAND_NAME"])))
        
        
        cent_off_types = sorted(list(set(promo_df['CENT_OFF_BY_PIECE_TYPE'])))
        
        ### FILTERS
        ## 1st LAYER
        brand_col, _, global_col = st.columns([3,0.2, 2])
        with brand_col :
            if user_action !=  "Review predicted promotions":
                brand_filter = st.selectbox(
                    "Brand",
                    brands
                )
                if not predicted_promo_df.empty:
                    predicted_promo_df = predicted_promo_df[predicted_promo_df["BRAND_NAME"] == brand_filter]
                    if brand_filter:
                        categories_df = categories_df[categories_df["BRAND_NAME"] == brand_filter]
            else: 
                brand_filter = st.multiselect(
                    "Brand",
                    brands,
                    []
                )
                if brand_filter :
                    if not predicted_promo_df.empty:
                        predicted_promo_df = predicted_promo_df[predicted_promo_df["BRAND_NAME"].isin(brand_filter)]
        
        # Define the lists of product categories depending on the selected brand(s)
        if user_action ==  "Review predicted promotions":
            if len(brand_filter) == 1:
                categories_df = categories_df[categories_df["BRAND_NAME"] == brand_filter[0]]
            else :
                categories_df = categories_df[categories_df["BRAND_NAME"].isin(['ANDROS', 'BONNE MAMAN'])]
        else:
            categories_df = categories_df[categories_df["BRAND_NAME"] == brand_filter]
                
        product_category = list(set(categories_df["KATEGORIE"]))
        product_category = sorted(product_category, key=str.casefold)
        
        product_undercategory = list(set(categories_df["UNTERKATEGORIE"]))
        product_undercategory = list(filter(lambda x: x is not None, product_undercategory)) 
        product_undercategory = sorted(product_undercategory, key=str.casefold)
        
        product_segment = list(set(categories_df["SEGMENT"]) )
        product_segment = list(filter(lambda x: x is not None, product_segment)) 
        product_segment = sorted(product_segment, key=str.casefold)
        
        product_undersegment = list(set(categories_df["UNTERSEGMENT"]))
        product_undersegment = list(filter(lambda x: x is not None, product_undersegment)) 
        product_undersegment = sorted(product_undersegment, key=str.casefold)
        
        categories_df['SKU_ID'] = categories_df.apply(lambda row: f"{row['SKU_BESCHREIBUNG']} ({row['CLIENT_ARTIKEL_ID']})" if row['CLIENT_ARTIKEL_ID'] else row['SKU_BESCHREIBUNG'], axis=1)
        product_id = set(categories_df["CLIENT_ARTIKEL_ID"])
        product_sku_id = list(set(categories_df["SKU_ID"]))
        product_sku_id = list(filter(lambda x: x is not None, product_sku_id)) 
        product_sku_id = sorted(product_sku_id, key=str.casefold)

        with global_col :
            if user_action ==  "Review predicted promotions":
                is_global = st.selectbox(
                    "All Brands?",
                    ["Any", "No(False)", "Yes(True)"]
                )
            else:
                is_global = st.selectbox(
                    "All Brands?",
                    ["No(False)", "Yes(True)"]
                )
        if is_global == "No(False)":
            is_global = False
            if not predicted_promo_df.empty:
                predicted_promo_df = predicted_promo_df[predicted_promo_df["IS_GLOBAL"] == is_global]
        elif is_global == "Yes(True)":
            is_global = True
            if not predicted_promo_df.empty:
                predicted_promo_df = predicted_promo_df[predicted_promo_df["IS_GLOBAL"] == is_global]
        
        ## 2nd LAYER
        year_col, _, coop_week_col, _, _ = st.columns([2,0.2, 3, 0.2, 3])
        with year_col :
            year_filter = st.selectbox(
                "Year",
                years
            )
        if year_filter != "All":
            if not predicted_promo_df.empty:
                predicted_promo_df = predicted_promo_df[predicted_promo_df["YEAR"] == year_filter]
        
        with coop_week_col :
            week_filter = st.multiselect(
                "Coop Week(s)",
                weeks,
                []
            )
        
        if len(week_filter) > 1 :
            weeks_str = [str(x) for x in week_filter]
            weeks_str_formatted = "+".join(weeks_str)
            if not predicted_promo_df.empty:
                predicted_promo_df = predicted_promo_df[ (predicted_promo_df["WEEK"] == weeks_str_formatted) |(promo_df["WEEK"].str.split('+').apply(lambda x: any(week in x for week in weeks_str)))]
        elif len(week_filter) == 1:
            weeks_str = str(week_filter[0])
            if not predicted_promo_df.empty:
                predicted_promo_df = predicted_promo_df[ (predicted_promo_df["WEEK"] == weeks_str) |(promo_df["WEEK"].str.split('+').apply(lambda x: any(week in x for week in weeks_str)))]
    
        ## 3rd LAYER
        promo_type_col, duration_type_col, duration_col, mechanism_col = st.columns([2, 2, 2, 3])
        with promo_type_col :
            promo_type_filter = st.selectbox(
                "Promo Type",
                promo_types
            )
        if promo_type_filter == "Other":
            with promo_type_col:
                promo_type_filter = st.text_input("New Promo Type:")
        if promo_type_filter != "Any":
            if not predicted_promo_df.empty:
                predicted_promo_df = predicted_promo_df[predicted_promo_df["PROMOTION_TYPE"] == promo_type_filter]
        
        with duration_type_col :
            duration_type_filter = st.selectbox(
                "Duration",
                duration_types
            )
        max_duration = 1
        if duration_type_filter == 'Woche(n)':
            max_duration = 52
        else: 
            max_duration = 7
        duration_filter_input=""
        if duration_type_filter != "Any":
            with duration_col :
                duration_filter = st.number_input(
                    "",
                    min_value=1,
                    max_value=max_duration
                )
                duration_filter_input = f"{str(duration_filter)} {duration_type_filter[:-3]}"
                if duration_filter > 1 :
                    if duration_type_filter == 'Woche(n)':
                        duration_filter_input += "n"
                    else:
                        duration_filter_input += "e" 
            duration_filter = f"{str(duration_filter)} {duration_type_filter[0]}"
            if not predicted_promo_df.empty:
                predicted_promo_df = predicted_promo_df[predicted_promo_df["DURATION"].str[:3] == duration_filter]
            
        with mechanism_col :
            mechanism_filter = st.selectbox(
                "Promo Mechanism",
                mechanisms
            )
        if mechanism_filter == "Other":
            with mechanism_col:
                mechanism_filter = st.text_input("New Promo Mechanism:")
        if mechanism_filter != "Any":
            if not predicted_promo_df.empty:
                predicted_promo_df = predicted_promo_df[predicted_promo_df["MECHANISM"] == mechanism_filter]
    
        ## 4th LAYER
        region_filter = st.multiselect(
            "Region(s)",
            regions,
            []
        )
        if region_filter:
            if not predicted_promo_df.empty:
                predicted_promo_df = predicted_promo_df[predicted_promo_df["REGION"].str.split('+').apply(lambda x: any(region in x for region in region_filter))]
        
        ## 5th LAYER
        cat_col, undercat_col, seg_col, underseg_col = st.columns([1, 1, 1, 1])
        with cat_col :
            cat_filter = st.multiselect(
                "Product Category(ies)",
                sorted(product_category)
        )
        if cat_filter:
            if not predicted_promo_df.empty:
                predicted_promo_df = predicted_promo_df[
                        predicted_promo_df["PRODUCT_CATEGORY"].apply(lambda x: any(cat in x.split(' && ') for cat in cat_filter) if x and x != "nan" else True)
                ]
            # filter available undercategories depending on chosen categories
            product_undercategory = sorted(list(set(categories_df[categories_df["KATEGORIE"].isin(cat_filter)]
                                        ["UNTERKATEGORIE"]
                                        .dropna()
                                        .unique()
                                        )))
            product_segment = sorted(list(set(categories_df[categories_df["KATEGORIE"].isin(cat_filter)]
                                ["SEGMENT"]
                                .dropna()
                                .unique()
                                )))
            product_undersegment = sorted(list(set(categories_df[categories_df["KATEGORIE"].isin(cat_filter)]
                                        ["UNTERSEGMENT"]
                                        .dropna()
                                        .unique()
                                        )))
            product_sku_id = sorted(list(set(categories_df[categories_df["KATEGORIE"].isin(cat_filter)]
                                ["SKU_ID"]
                                .dropna()
                                .unique()
                                )))
        
        else : 
            cat_filter = sorted(list(product_category))
                
        with undercat_col :
            undercat_filter = st.multiselect(
                "Undercategory(ies)",
                product_undercategory
        )
            
        if undercat_filter:
            if not predicted_promo_df.empty:
                predicted_promo_df = predicted_promo_df[
                    predicted_promo_df["PRODUCT_UNDERCATEGORY"].apply(lambda x: any(undercat in x.split(' && ') for undercat in undercat_filter) if x and x != "nan" else True)
                ]
            
            
            # filter available undercategories depending on chosen categories
            product_segment = sorted(list(set(categories_df[
                                    (categories_df["UNTERKATEGORIE"].isin(undercat_filter))
                                    ]
                                    ["SEGMENT"]
                                    .dropna()
                                    .unique()
                                 )))
            product_undersegment = sorted(list(set(categories_df[
                                       (categories_df["UNTERKATEGORIE"].isin(undercat_filter))
                                    ]
                                    ["UNTERSEGMENT"]
                                    .dropna()
                                    .unique()
                                 )))
            product_sku_id = sorted(list(set(categories_df[ 
                                (categories_df["UNTERKATEGORIE"].isin(undercat_filter))
                                ]
                                ["SKU_ID"]
                                .dropna()
                                .unique()
                                )))
    
        else : 
            if user_action ==  "Review predicted promotions":
                undercat_filter = sorted(list(product_undercategory))
        with seg_col :
            seg_filter = st.multiselect(
                "Segment(s)",
                product_segment
        )
        if seg_filter:
            if not predicted_promo_df.empty:
                predicted_promo_df = predicted_promo_df[
                    predicted_promo_df["PRODUCT_SEGMENT"].apply(lambda x: any(seg in x.split(' && ') for seg in seg_filter) if x and x != "nan" else True)
                ]
    
            # filter available undercategories depending on chosen categories
            product_undersegment = sorted(list(set(categories_df[
                                    (categories_df["SEGMENT"].isin(seg_filter))
                                    ]
                                    ["UNTERSEGMENT"]
                                    .dropna()
                                    .unique()
                                    )))
            product_sku_id = sorted(list(set(categories_df[
                                (categories_df["SEGMENT"].isin(seg_filter))
                                ]
                                ["SKU_ID"]
                                .dropna()
                                .unique()
                                )))
        else : 
            if user_action ==  "Review predicted promotions":
                seg_filter = list(product_segment)
            
        with underseg_col :
            underseg_filter = st.multiselect(
                "Undersegment(s)",
                product_undersegment
        )
        if underseg_filter:
            if not predicted_promo_df.empty:
                predicted_promo_df = predicted_promo_df[
                    predicted_promo_df["PRODUCT_UNDERSEGMENT"].apply(lambda x: any(underseg in x.split(' && ') for underseg in underseg_filter) if x and x != "nan" else True)   
                ]
            # filter available undercategories depending on chosen categories
            product_sku_id = sorted(list(set(categories_df[
                                (categories_df["UNTERSEGMENT"].isin(underseg_filter))
                                ]
                                ["SKU_ID"]
                                .dropna()
                                .unique()
                                )))
        else : 
            if user_action ==  "Review predicted promotions":
                underseg_filter = list(product_undersegment)
            
        ## 6th LAYER
        product_id_filter = st.multiselect(
            "Product SKU(s)",
            product_sku_id
        ) 
        if product_id_filter:
            product_id_filter_formatted=[]
            for i in range(0, len(product_id_filter)):
                product_id = product_id_filter[i][-10:-1]
                product_sku = product_id_filter[i][0:-12]
                product_id_formatted = str(product_id) + " " + str(product_sku)
                product_id_filter_formatted.append(product_id_formatted)
            predicted_promo_df = predicted_promo_df[
                predicted_promo_df["PRODUCT_ID"].apply(lambda x: any(id in x.split(' && ') for id in product_id_filter_formatted) if x and x != "nan" else True)   
            ]
                    
    
        ## Values Input
        st.write("---")
        if user_action !=  "Review predicted promotions":
            WB_col, cent_off_type_col, cent_off_col, promoted_perc_col = st.columns([1, 1, 1, 1])
            with WB_col:
                WB_filter = st.number_input(
                    "WB (CHF)",
                    min_value=0.00,
                    step=1.00
                )
            with cent_off_type_col:
                cent_off_type_filter = st.selectbox(
                    "Cent Off By Piece Type",
                    cent_off_types
                )
            with cent_off_col:
                cent_off_filter = st.number_input(
                    f"""Cent Off By Piece ({'CHF' if cent_off_type_filter == "AMOUNT" else '%'})""",
                    min_value=0.00,
                    step=1.00,
                    help="This represents the percentage of the price the products are sold to Coop" 
                            if cent_off_type_filter == "PERCENT" 
                            else "This represents the price at which the products are sold to Coop" 
                )
            with promoted_perc_col:
                promoted_vol_perc = st.number_input(
                    f"Promoted volume (%)",
                    value= 100.00,
                    min_value=0.00,
                    max_value=100.00,
                    step=1.00,
                    help="Sometimes your promotion mechanism means some of the promoted products will still be sold at the non promoted price" 
                )
                
                    

        ############
        ###### FORMAT VALUES ######
            
        ###### COOP WEEK ######
        input_weeks_str_formatted=""
        if len(week_filter) >= 1:
            for i in range(0, len(week_filter)):
                input_weeks_str_formatted += str(week_filter[i]) + "+"
            input_weeks_str_formatted = input_weeks_str_formatted[:-1]
        elif len(week_filter) == 0:
            input_weeks_str_formatted=None
        
        input_regions_str_formatted=""
        if len(region_filter) >= 1:
            for i in range(0, len(region_filter)):
                input_regions_str_formatted += str(region_filter[i]) + "+"
            input_regions_str_formatted = input_regions_str_formatted[:-1]
        else:
            input_regions_str_formatted=None
        
        
        ###### PRODUCT_CATEGORY ######
        input_categories_str_formatted=""
        if len(cat_filter) >= 1:
            for i in range(0, len(cat_filter)):
                input_categories_str_formatted += str(cat_filter[i]) + " && "
            input_categories_str_formatted = input_categories_str_formatted[:-4]     
        else:
            input_categories_str_formatted=None
        
            
        ###### PRODUCT_UNDERCATEGORY ######
        input_undercategories_str_formatted=""
        if len(undercat_filter) >= 1:
            for i in range(0, len(undercat_filter)):
                input_undercategories_str_formatted += str(undercat_filter[i]) + " && "
            input_undercategories_str_formatted = input_undercategories_str_formatted[:-4]     
        else:
            input_undercategories_str_formatted=None
    
        
        ###### PRODUCT_SEGMENT ######
        input_segment_str_formatted=""
        if len(seg_filter) >= 1:
            for i in range(0, len(seg_filter)):
                input_segment_str_formatted += str(seg_filter[i]) + " && "
            input_segment_str_formatted = input_segment_str_formatted[:-4]     
        else:
            input_segment_str_formatted=None
    
        
        ###### PRODUCT_UNDERSEGMENT ######
        input_undersegment_str_formatted=""
        if len(underseg_filter) >= 1:
            for i in range(0, len(underseg_filter)):
                input_undersegment_str_formatted += str(underseg_filter[i]) + " && "
            input_undersegment_str_formatted = input_undersegment_str_formatted[:-4]     
        else:
            input_undersegment_str_formatted=None
        
        
        ###### PRODUCT_ID ######
        input_product_id_str_formatted=""
        if len(product_id_filter) >= 1:
            for i in range(0, len(product_id_filter)):
                product_id = product_id_filter[i][-10:-1]
                product_sku = product_id_filter[i][0:-12]
                input_product_id_str_formatted += str(product_id) + " " + str(product_sku) + " && "
            input_product_id_str_formatted = input_product_id_str_formatted[:-4]  
            #st.write(input_product_id_str_formatted)
        else:
            input_product_id_str_formatted=None
            
        if user_action !=  "Review predicted promotions":
            ##### CREATE NEW ROW #####
            promo_level = ""
            new_promo_df = pd.DataFrame({
                        'YEAR':[year_filter],
                        'WEEK':[input_weeks_str_formatted], 
                        'COOP_WEEK':[input_weeks_str_formatted],
                        'PROMOTION_TYPE':[promo_type_filter],
                        'IS_GLOBAL':[is_global],
                        'DURATION':[duration_filter_input],
                        'REGION':[input_regions_str_formatted],
                        'MECHANISM':[mechanism_filter],
                        'BRAND_NAME':[brand_filter],
                        'PROMOTION_LEVEL' :[promo_level],
                        'PRODUCT_CATEGORY':[input_categories_str_formatted if input_categories_str_formatted != "" else None],
                        'PRODUCT_UNDERCATEGORY':[input_undercategories_str_formatted if input_undercategories_str_formatted != "" else None],
                        'PRODUCT_SEGMENT':[input_segment_str_formatted if input_segment_str_formatted != "" else None],
                        'PRODUCT_UNDERSEGMENT':[input_undersegment_str_formatted if input_undersegment_str_formatted != "" else None],
                        'PRODUCT_ID':[input_product_id_str_formatted if input_product_id_str_formatted != "" else None],
                        'WB':[WB_filter],
                        'CENT_OFF_BY_PIECE':[cent_off_filter],
                        'CENT_OFF_BY_PIECE_TYPE':[cent_off_type_filter],
                        'PROMOTED_VOLUME_PERC':[promoted_vol_perc],
                        'TIMESTAMP' : [datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
                        'USER_ID' : [st.experimental_user.email]
                    })
            level_df = new_promo_df[["PRODUCT_CATEGORY", "PRODUCT_UNDERCATEGORY", "PRODUCT_SEGMENT", "PRODUCT_UNDERSEGMENT", "PRODUCT_ID"]].copy()
    
            ##### CALCULATE PROMO_LEVEL #####
            promo_level_df = level_df.copy() 
            nan_value = float("NaN")
            promo_level_df.replace("", nan_value, inplace=True)
            promo_level_df.dropna(how='all', axis=1, inplace=True)
            if len(promo_level_df.columns) == 0:
                promo_level_col = ""
            if promo_level_df.empty:
                st.warning("Please chose at least one category, undercategory, segment AND/OR undersegment.")
            else:
                promo_level_col = promo_level_df.columns[-1]
                if promo_level_col == "PRODUCT_ID":
                    promo_level = "PRODUCT"
                else:
                    promo_level = promo_level_col.split('_', 1)[1]
                new_promo_df["PROMOTION_LEVEL"]=promo_level
        
        
                ##### CALCULATE ALL LEVELS #####
                levels = list(level_df.columns)
                last_level_index = levels.index(promo_level_col)
                last_level_value = new_promo_df[promo_level_col][0]
                last_level_values = last_level_value.split(" && ")
                if promo_level_col == "PRODUCT_ID":
                    category_level_col = "CLIENT_ARTIKEL_ID"
                    for i in range (0, len(last_level_values)):
                        last_level_values[i] = last_level_values[i][:9]
                        
                else: 
                    category_level_col = promo_level.replace("D", "T")
                    category_level_col = category_level_col.replace("C", "K")
                    category_level_col = category_level_col.replace("Y", "IE")
        
                for i in range (1, last_level_index+1):
                    curr_level_col = levels[last_level_index - i]
                    curr_cat_col = curr_level_col
                    curr_cat_col = curr_cat_col.split('_', 1)[1]
                    curr_cat_col = curr_cat_col.replace("D", "T")
                    curr_cat_col = curr_cat_col.replace("C", "K")
                    curr_cat_col = curr_cat_col.replace("Y", "IE")
                        
                    curr_level_value = categories_df[(categories_df[category_level_col].isin(last_level_values))][curr_cat_col].reset_index(drop=True)
                    curr_level_value = " && ".join(sorted(set(filter(None, curr_level_value))))
    
                    new_promo_df[curr_level_col]=curr_level_value if curr_level_value != "" else None
                        
                    product_sku_id = set(categories_df[
                                            (categories_df["UNTERSEGMENT"].isin(underseg_filter))
                                            ]
                                            ["SKU_ID"]
                                            )
                st.write("---")
                
                st.write("""
                            Here is the resulting promotion below. \n\n
                            Please check that all cells are filled correctly. 
                            If there is a mistake, please correct it using the option boxes above. \n\n
                            Once you are sure about your data, you can predict the impact of this promotion by pressing the button below.
                            """
                            )

                column_config = {
                    "IS_GLOBAL": st.column_config.Column(
                        "ALL BRANDS"
                    )
                }
                result_df = st.data_editor(
                                new_promo_df, 
                                disabled=new_promo_df.columns, 
                                column_order=('YEAR', 'COOP_WEEK', 'PROMOTION_TYPE', 'IS_GLOBAL', 'DURATION',
                                            'REGION', 'MECHANISM', 'BRAND_NAME', 
                                            'PRODUCT_CATEGORY', 'PRODUCT_UNDERCATEGORY', 'PRODUCT_SEGMENT', 
                                            'PRODUCT_UNDERSEGMENT', 'PRODUCT_ID', 'WB', 'CENT_OFF_BY_PIECE', 
                                            'CENT_OFF_BY_PIECE_TYPE', 'PROMOTED_VOLUME_PERC', 'TIMESTAMP', 'USER_ID'),                                
                                hide_index = True, 
                                column_config=column_config,
                                width=10000
                )
               
                missing_vals_count = 0
                mandatory_fileds_str = ""
                mandatory_fileds = ['PROMOTION_TYPE', 'COOP_WEEK', 'REGION', 'MECHANISM']
                for i in range (0, len(mandatory_fileds)):
                    if result_df[mandatory_fileds[i]][0] == None or result_df[mandatory_fileds[i]][0] == "":
                            mandatory_fileds_str += f"- {mandatory_fileds[i]} \n\n"
                            missing_vals_count += 1
                if missing_vals_count != 0 :
                    with st.container(border=True):
                        st.write("Careful! there are some empty values that should be filled, see the list below and make sure to fill them before saving. \n\nOtherwise you won't be able to predict your promotion impact.")
                        st.write("Missing values in columns: ")
                        st.write(mandatory_fileds_str)
                else:
                    save_button_col, save_text_col = st.columns((2, 3))
        
                    if 'was_saved' not in st.session_state:
                        st.session_state.was_saved = False
                        
                    loaded_table = SESSION.table(CONFIG["PROMO_FORECAST_TABLE"])
                    og_df = loaded_table.to_pandas()
                    with save_button_col:
                                # Button to save the changes made to the dataframe
                        if st.button("Predict promotion impact", key="save_test_partner", type='primary'):
                            st.session_state.missing_uptodate = False
                            st.session_state.was_saved = True
                    
                            with st.spinner('Your KPIs are being predicted'):
                                ### Add new row to BASELINE_REF_PROMO_FORECAST
                                update_table_no_truncate(
                                    CONFIG["PROMO_FORECAST_TABLE"], 
                                    new_promo_df, 
                                    numeric_columns=["PROMO_ID", "YEAR", "WB", "CENT_OFF_BY_PIECE", "PROMOTED_VOLUME_PERC", "NON_PROMOTED_VOLUME_PERC"]
                                )
                                
                            with save_text_col:
                                st.success('Your promotion was saved and is being predicted.They will appear below in a few seconds.')
                    
                    if st.session_state.was_saved == True:
                        with st.spinner('Your KPIs are being predicted'):
                            st.write("---")
                            loaded_table = SESSION.table(CONFIG["PROMO_FORECAST_TABLE"])
                            new_df = loaded_table.to_pandas()
                            new_rows_df = pd.concat([og_df, new_df]).drop_duplicates(keep=False)
                            new_rows_df = new_rows_df.drop_duplicates(subset=["FORECAST_PROMOTION_ID"], keep="first")
                            new_id = set(new_rows_df["FORECAST_PROMOTION_ID"])
                            prediction_table = SESSION.table(CONFIG["PROMO_FORECAST_KPI_TABLE"])
                            predicted_promo_df = prediction_table.to_pandas()
                            predicted_row_df = predicted_promo_df[predicted_promo_df['FORECAST_PROMOTION_ID'].isin(new_id)]   

                            if not predicted_row_df.empty :
                                st.subheader("Here are your promotion impact KPIs")
                                with st.container(border=True):
                                    predicted_row_df = predicted_row_df.reset_index()
                                    unknown_val_txt = "Unknown value"
    
                                    # Retrieve KPI values and format them correctly for display
                                    baseline = round(predicted_row_df["BASELINE"][0])
                                    incremental_vol = round(predicted_row_df["INCREMENTAL_VOLUME"][0])
                                    incremental_margin = round(predicted_row_df["INCREMENTAL_MARGIN"][0])
                                    promoted_vol_perc = round(predicted_row_df["PROMOTED_VOLUME_PERC"][0])
                                    uplift_perc = round(predicted_row_df['UPLIFT'][0], 1)
                                    efficiency_perc = round(predicted_row_df['EFFICIENCY'][0], 1)
                                    
                                    total_vol = baseline + incremental_vol
    
                                    promoted_volume = promoted_vol_perc * total_vol * 0.01
                                    non_promoted_volume = total_vol - promoted_volume
                                    
                                    if total_vol != total_vol: 
                                        total_vol = f"""  *{unknown_val_txt}*"""
                                    else:
                                        total_vol = '{:,.0f}'.format(total_vol)
                                        total_vol = f""" #### {total_vol} """
                                        
                                    # if baseline != baseline: 
                                    #     baseline = f"""  *{unknown_val_txt}*"""
                                    # else:
                                    #     baseline = '{:,.2f}'.format(baseline)
                                    #     baseline = f"""#### {baseline} """
                                    # baseline_confidence_lvl = predicted_row_df["BASELINE_CONFIDENCE_LEVEL"][0]
                                    # if baseline_confidence_lvl != baseline_confidence_lvl: 
                                    #     baseline_confidence_lvl = f"""  *{unknown_val_txt}*"""
                                    # else:
                                    #     if baseline_confidence_lvl == 1:
                                    #         baseline_confidence_lvl = f""" :red[{baseline_confidence_lvl}] """
                                    #     elif baseline_confidence_lvl == 2:
                                    #         baseline_confidence_lvl = f""" :orange[{baseline_confidence_lvl}] """
                                    #     elif baseline_confidence_lvl == 3:
                                    #         baseline_confidence_lvl = f""" :green[{baseline_confidence_lvl}] """
                                    #     else:
                                    #         baseline_confidence_lvl = f""" :black[{baseline_confidence_lvl}] """
                                    
                                    
                                    if incremental_vol != incremental_vol: 
                                        incremental_vol = f"""  *{unknown_val_txt}*"""
                                    else:
                                        incremental_vol = '{:,.0f}'.format(incremental_vol)
                                        incremental_vol = f""" #### {incremental_vol} """

                                    if uplift_perc != uplift_perc: 
                                        uplift_str = f"""  *{unknown_val_txt}*"""
                                    else:
                                        uplift_str = '{:,.1f}'.format(uplift_perc)
                                        uplift_str = f""" #### {uplift_str} """

                                    if efficiency_perc != efficiency_perc: 
                                        efficiency_str = f"""  *{unknown_val_txt}*"""
                                    else:
                                        efficiency_str = '{:,.1f}'.format(efficiency_perc)
                                        efficiency_str = f""" #### {efficiency_str} """

    
                                    if incremental_margin != incremental_margin: 
                                        incremental_margin = f"""  *{unknown_val_txt}*"""
                                    else:
                                        incremental_margin = '{:,.0f}'.format(incremental_margin)
                                        incremental_margin = f""" #### {incremental_margin} """
                                    
                                    # if promoted_volume != promoted_volume: 
                                    #     promoted_volume = f"""  *{unknown_val_txt}*"""
                                    # else:
                                    #     promoted_volume = '{:,.0f}'.format(promoted_volume)
                                    #     promoted_volume = f""" {promoted_volume} """
    
                                    # if non_promoted_volume != non_promoted_volume: 
                                    #     non_promoted_volume = f"""  *{unknown_val_txt}*"""
                                    # else:
                                    #     non_promoted_volume = '{:,.2f}'.format(non_promoted_volume)
                                    #     non_promoted_volume = f""" {non_promoted_volume} """
                                    
                                    investment = round(predicted_row_df["INVESTMENT_TOTAL"][0])
                                    if investment != investment: 
                                        investment_str = f"""  *{unknown_val_txt}*"""
                                    else:
                                        tot_investment = '{:,.0f}'.format(investment)
                                        investment_str = f""" {tot_investment} """
                                    roi_percent = round(predicted_row_df["ROI_PERCENT"][0], 1)
                                    if roi_percent != roi_percent: 
                                        roi_percent = f"""  *{unknown_val_txt}*"""
                                    else:
                                        roi_percent = f""" {roi_percent}% """
                                    fix_investment = round(predicted_row_df['INVESTMENT_FIXED'][0])
                                    fix_investment_str = '{:,.0f}'.format(fix_investment)
                                    var_investment = round(predicted_row_df['INVESTMENT_VARIABLE'][0])
                                    var_investment_str = '{:,.0f}'.format(var_investment)
                                    promoimpact_confidence_lvl = predicted_row_df["PROMOIMPACT_CONFIDENCE_LEVEL"][0]
                                    if promoimpact_confidence_lvl != promoimpact_confidence_lvl: 
                                        promoimpact_confidence_lvl = f"""  *{unknown_val_txt}*"""
                                    else:
                                        if promoimpact_confidence_lvl == 1:
                                            promoimpact_confidence_lvl = f""" :red[{promoimpact_confidence_lvl}] """
                                        elif promoimpact_confidence_lvl == 2:
                                            promoimpact_confidence_lvl = f""" :orange[{promoimpact_confidence_lvl}] """
                                        elif promoimpact_confidence_lvl == 3:
                                            promoimpact_confidence_lvl = f""" :green[{promoimpact_confidence_lvl}] """
                                        else:
                                            promoimpact_confidence_lvl = f""" :black[{promoimpact_confidence_lvl}] """
                                        
                                    volumes_col, investment_col = st.columns([1, 1])
                                    with volumes_col:
                                        with st.container(border=True):
                                            st.markdown(f""" 
                                                #### VOLUMES
                                            """)
                                            st.write("")
                                            # st.markdown(f""" 
                                            #     ###### Baseline 
                                            #     \n\n{baseline}p
                                            #     \n\n Confidence level: {baseline_confidence_lvl}
                                            # """)
                                            st.write("")
                                            st.markdown(f""" 
                                                \n\n ###### Total Volume
                                                \n\n {total_vol}p
                                            """)
                                            # promo_col, non_promo_col = st.columns([1, 1.1])
                                            # with promo_col : 
                                            #     st.markdown(f""" 
                                            #          **Promoted Volume:** 
                                            #          {promoted_volume}p
                                            #     """)
                                            # with non_promo_col : 
                                            #     st.markdown(f""" 
                                            #          **Non Promoted Volume:** 
                                            #          {non_promoted_volume}p
                                            #     """)
                                            st.write("")
                                            st.markdown(f""" 
                                                ###### Incremental Volume
                                                \n\n {incremental_vol}p
                                                
                                            """)
                                            st.write("")
                                            st.markdown(f""" 
                                                \n\n ###### Uplift
                                                \n\n {uplift_str}%
                                            """)

                                            st.write("")
                                            st.markdown(f""" 
                                                \n\n ###### Efficiency
                                                \n\n {efficiency_str}%
                                            """)
                                            
                                            
                                    with investment_col:
                                        with st.container(border=True):
                                            st.markdown(f""" 
                                                #### INVESTMENT
                                            """)   
                                            st.write("")
                                            st.markdown(f""" 
                                                ###### Total Investment
                                                \n\n #### {investment_str} CHF
                                            """)
                                            st.write("")
                                            fix_col, var_col = st.columns([1, 1.1])
                                            with fix_col : 
                                                st.markdown(f""" 
                                                     **Fixed Investment:** 
                                                     \n\n**{fix_investment_str} CHF**
                                                """)
                                            with var_col : 
                                                st.markdown(f""" 
                                                     **Variable Investment:** 
                                                     \n\n**{var_investment_str} CHF**
                                                """)      
                                            st.write("")
                                            st.markdown(f""" 
                                                \n\n ###### Incremental Margin
                                                \n\n {incremental_margin} CHF
                                            """)
                                            st.write("")    
                                            st.markdown(f""" 
                                                ###### ROI
                                                \n\n #### {roi_percent} 
                                            """)
                                    
                                                                        # DÉBUT DE LA MODIFICATION
                                    with st.container(border=True, height=120):
                                        # Utilisation de st.columns pour aligner le texte du titre et le popover
                                        # Ajustez les ratios si l'espacement n'est pas parfait
                                        title_col, info_popover_col, _ = st.columns([0.15, 0.03, 0.82]) # Je suggère 0.7, 0.3 pour plus de place pour le titre
                                        
                                        with title_col:
                                            st.markdown(f""" 
                                                ###### Promotion Impact Confidence level
                                            """)
                                        
                                        with info_popover_col:
                                            # Utilise st.popover pour afficher le texte au clic
                                            with st.popover(
                                                "ℹ️", # Label du bouton du popover (l'émoji)
                                                use_container_width=True,
                                                # key="promo_confidence_info_popover" # Clé unique
                                            ):
                                                # Le contenu de la bulle d'aide est défini ici (Markdown supporté)
                                                st.markdown(
                                                    "The calculation of the promotion's impact is based on a growth rate, allowing for the measurement of the total products sold, compared to a baseline without promotion. According to the formula growth rate = 100 * ([Total Sales] - [Baseline]) / [Baseline] This growth rate is calculated from the history of Andros & Bonne Maman promotions and varies according to the type of promotion and the products concerned.\n\n"
                                                    "The confidence index evaluates the reliability of the growth rate and therefore the impact of the promotion. This confidence index has 3 levels:\n\n"
                                                    "**Level 1 (Limited Confidence):** There are fewer than 5 promotions similar to the predicted promotion in the database\n\n"
                                                    "**Level 2 (Medium Confidence):** There are at least 5 promotions similar to the predicted promotion in the database, but the variability of their results is high\n\n"
                                                    "**Level 3 (High Confidence):** There are at least 5 promotions similar to the predicted promotion in the database and their results are stable\n\n"
                                                                                                    )

                                        st.markdown(f""" #### {promoimpact_confidence_lvl} """)
                                    # FIN DE LA MODIFICATION
                                    
        else:
            st.subheader('Predicted promotions')
            if predicted_promo_df.empty:
                st.warning("You haven't predicted any promotions yet")
            else:
                df_sorted = predicted_promo_df.copy()
                def get_min_week(week_str):
                    if pd.isna(week_str):
                        return float('inf')
                    return min(int(w) for w in week_str.split('+'))
                
                df_sorted['WEEK_MIN'] = df_sorted['COOP_WEEK'].apply(get_min_week)
                df_sorted = df_sorted.sort_values(['YEAR', 'WEEK_MIN'], ascending=[True, True])
                column_config = {
                    "IS_GLOBAL": st.column_config.Column(
                        "ALL BRANDS"
                    )
                }
                df_sorted = df_sorted.drop('WEEK_MIN', axis=1)

                # Modification 1 : Arrondir les nombres dans le DataFrame pour l'affichage (ton code existant)
                df_to_display = df_sorted.copy()
                numeric_cols = df_to_display.select_dtypes(include='number').columns
                df_to_display[numeric_cols] = df_to_display[numeric_cols].round(0).astype('Int64')

                column_config = {
                    "IS_GLOBAL": st.column_config.Column(
                        "ALL BRANDS"
                    ),
                    # Assure-toi que PROMOTION_ID est désactivé même s'il est une colonne principale
                    "FORECAST_PROMOTION_ID": st.column_config.NumberColumn(
                        label="Promotion ID",
                        help="Unique identifier for the promotion",
                        disabled=True # Très important pour empêcher l'édition de l'ID
                    )
                }


                # Stocke les IDs originaux pour la comparaison après l'édition
                original_promo_ids = set(df_to_display["FORECAST_PROMOTION_ID"])

                # Utilise st.data_editor pour permettre la suppression de lignes
                predicted_promo_df_edited = st.data_editor(
                    df_to_display,
                    hide_index=True,
                    column_config=column_config, # Utilise les configurations de colonnes existantes
                    width=10000,
                    num_rows='dynamic',  # Empêche l'ajout de nouvelles lignes
                    disabled=df_to_display.columns.tolist() # Désactive toutes les colonnes pour l'édition de valeurs
                )
                # Vérifie les lignes supprimées par l'utilisateur
                current_promo_ids = set(predicted_promo_df_edited["FORECAST_PROMOTION_ID"])
                deleted_ids = original_promo_ids - current_promo_ids # IDs qui étaient là mais ne le sont plus


                save_as_excel_button(predicted_promo_df, "predicted_promotions")
                # Ajoute un bouton "Sauvegarder les suppressions"
                save_button_col, save_text_col = st.columns((2, 3))
                with save_button_col:
                    if st.button("Save Deletions", key="save_predictions_deletions", type="primary"):
                        if len(deleted_ids) > 0:
                            with st.spinner('Deleting selected promotions...'):
                                # Appel de ta fonction delete_rows_table
                                # IMPORTANT : Utilise la table correcte pour les prédictions (KPI_TABLE)
                                delete_rows_table(CONFIG["PROMO_FORECAST_TABLE"], list(deleted_ids))
                            
                            with save_text_col :
                                st.success(f"Successfully deleted {len(deleted_ids)} promotion(s).")
                                # Définit was_saved à True pour recharger les données après la suppression
                                st.session_state.was_saved = True
                                st.rerun()
                        else:
                            with save_text_col :
                                st.info("No promotions selected for deletion.")
                        
                tot_promo_nb = len(predicted_promo_df.index)
                
                # MODIFICATION 2 : Changer le formatage de .2f à .0f pour tous les totaux
                
                # tot_baseline_vol = predicted_promo_df['BASELINE'].sum()
                # tot_baseline_vol_str = '{:,.0f}'.format(tot_baseline_vol)
                
                tot_increment_vol = round(predicted_promo_df['INCREMENTAL_VOLUME'].sum())
                tot_increment_vol_str = '{:,.0f}'.format(tot_increment_vol)
                
                tot_vol = round(predicted_promo_df['TOTAL_VOLUME'].sum())
                tot_vol_str = '{:,.0f}'.format(tot_vol)

                uplift = round(tot_increment_vol/(predicted_promo_df['BASELINE'].sum())*100, 1)
                uplift_str = '{:,.1f}'.format(uplift)

                efficiency = round(tot_increment_vol/(predicted_promo_df['PROMOTED_VOLUME'].sum())*100, 1)
                efficiency_str = '{:,.1f}'.format(efficiency)
                
                tot_increment_val = round(predicted_promo_df['INCREMENTAL_MARGIN'].sum())
                tot_increment_val_str= '{:,.0f}'.format(tot_increment_val)
                
                # promoted_vol = predicted_promo_df['PROMOTED_VOLUME'].sum()
                # promoted_vol_str = '{:,.0f}'.format(promoted_vol)
                
                tot_investment = round(predicted_promo_df['INVESTMENT_TOTAL'].sum())
                tot_investment_str = '{:,.0f}'.format(tot_investment)
                
                fix_investment = predicted_promo_df['INVESTMENT_FIXED'].sum()
                fix_investment_str = '{:,.0f}'.format(fix_investment)
                
                var_investment = predicted_promo_df['INVESTMENT_VARIABLE'].sum()
                var_investment_str = '{:,.0f}'.format(var_investment)
                
                avg_roi = round((tot_increment_val - tot_investment) / tot_investment * 100, 1) if tot_investment != 0 else 0
                # Pour le ROI, vous voudrez peut-être garder des décimales, mais si non, voici comment l'arrondir :
                avg_roi_str = '{:,.1f}'.format(avg_roi)
                
                st.write("")
                st.write("")
                with st.container(border=True):
                    st.subheader('Summary')
                    st.markdown(f""" 
                                ##### Total number of promotions: {tot_promo_nb}
                            
                            """)
                    left_col, right_col = st.columns([1, 1])
                    with left_col:
                        
                        with st.container(border=True):
                            st.markdown(f""" #### VOLUMES """)
                            st.write("")
                            # st.markdown(f""" 
                            #     \n\n ##### Total Baseline:
                            #     \n\n ### {tot_baseline_vol_str} p 
                            
                            # """)
                            
                            # st.write("")
                            # st.markdown(f""" 
                            #     \n\n ##### Total Incremental:
                            #     \n\n ### {tot_increment_vol_str} p 
                            
                            # """)
                            # st.write("")
                            # st.markdown(f""" 
                            #     \n\n ##### Total Volume:
                            #     \n\n ### {tot_vol_str} p 
                            
                            # """)
                            # st.write("")
                            # st.markdown(f""" 
                            #     \n\n ##### Total Promoted volume:
                            #     \n\n ### {promoted_vol_str} p
                            st.markdown(f""" 
                                \n\n ##### Total Volume:
                                \n\n ### {tot_vol_str} p 
                            
                            """)
                            
                            st.write("")
                            st.markdown(f""" 
                                \n\n ##### Incremental volume:
                                \n\n ### {tot_increment_vol_str} p 
                            
                            """)
                            st.write("")
                            st.markdown(f""" 
                                \n\n ##### Uplift:
                                \n\n ### {uplift_str} %
                            
                            """)
                            st.write("")
                            st.markdown(f""" 
                                \n\n ##### Efficiency:
                                \n\n ### {efficiency_str} % 
                            
                            """)
                            
                    with right_col:
                        
                        with st.container(border=True):
                            st.markdown(f""" #### INVESTMENT""")

                            st.write("")
                            st.markdown(f""" 
                                \n\n ##### Total Investment:
                                \n\n ### {tot_investment_str} CHF
                            """)
                            st.write("")
                            fix_col, var_col = st.columns([1, 1.1])
                            with fix_col : 
                                st.markdown(f""" 
                                        **Fixed Investment:** 
                                        \n\n**{fix_investment_str} CHF**
                                """)
                            with var_col : 
                                st.markdown(f""" 
                                        **Variable Investment:** 
                                        \n\n**{var_investment_str} CHF**
                                """)      
            
                            st.write("")
                            st.markdown(f""" 
                                    \n\n ##### Incremental margin :  
                                    \n\n ### {tot_increment_val_str} CHF
                                
                                """)
                            
                            st.write("")
                            st.markdown(f""" 
                                    \n\n ##### ROI:
                                    \n\n ### {avg_roi_str} %
                                
                                """)
                        
 
if __name__ == '__main__':
    
    #The main function that initializes the Streamlit application. 
    #It retrieves the user's role, and calls the body of the application
    #to display content based on the user's role.
    
    #Returns:
    #- None
    

    # Call the body function to render the main content of the app, passing in the user's role
    body()



    
def filter_df(self, cols, vals):
        """
        Filters the DataFrame based on matching column values. Returns a new instance of 
        StreamlitDataframe containing the filtered DataFrame.

        Args:
            cols (list): A list of column names to filter by.
            vals (list): A list of values corresponding to each column in 'cols'.

        Returns:
            StreamlitDataframe: A new StreamlitDataframe instance with the filtered DataFrame, 
                                or None if the number of columns and values do not match.
        """

        # Check if the number of columns matches the number of values
        if len(cols) != len(vals): 
            st.warning('Your number of columns if different from your number of values')
            return None
        else: 
            # Filter the DataFrame based on the specified columns and values
            filtered_df = self.df.copy()

            for i_col in range (0, len(cols)):
                if vals[i_col] != 'All':
                    if isinstance(vals[i_col], str):
                        filtered_df = filtered_df[filtered_df[f'{cols[i_col]}'] == vals[i_col]]
                    else:
                        filtered_df = filtered_df[filtered_df[f'{cols[i_col]}'] == vals[i_col]]            
                
            # Return a new StreamlitDataframe instance containing the filtered DataFrame
            return filtered_df
            
def save_as_csv_button(df, df_name):
        """
        This function generates a Streamlit button to allow users to download a DataFrame as a CSV file.

        - The DataFrame (`df`) is first rounded to 2 decimal places for consistency in numerical data.
        - The `convert_df` function, which is decorated with `@st.cache_data`, converts the DataFrame to a CSV format 
        and caches the result to avoid redundant computations during reruns.
        - The CSV data is then passed to the `st.download_button` method, which creates a download button in the Streamlit app.
        - When the button is clicked, the CSV file will be downloaded with the specified file name (`df_name.csv`).
        """
        # Round all numbers so they only have two numbers after the comma
        df = df.round(2)

        @st.cache_data
        def convert_df(df):
            # Cache the DataFrame conversion to CSV format to improve efficiency on reruns
            return df.to_csv().encode("utf-8")
        
        csv = convert_df(df)
        
        st.download_button(
            label="Download data as CSV",
            data=csv,
            file_name=f"{df_name}.csv",
            mime="text/csv",
        )
        



