import os
import re
import pandas as pd
import docx
import win32com.client as win32
from win32com.client import constants

def clean_text(text):
    """Cleans up text by stripping whitespace and removing common artifacts."""
    if pd.isna(text):  # Handle NaN values
        return ""
    text = str(text) # Ensure text is a string
    text = re.sub(r'\s+', ' ', text).strip()
    text = text.replace("_x000D_", " ") # Replace Word's newline artifact
    text = text.replace("\t", " ") # Replace tabs
    text = text.replace("\n", " ") # Replace newlines
    text = text.replace("x", "") # Remove 'x' often used for checkboxes
    return text.strip()

def clean_column_name(name):
    """Cleans and standardizes a column name, and attempts to shorten it."""
    name = clean_text(name)
    name = name.lower()
    name = name.replace(' ', '_')
    name = re.sub(r'[^a-zA-Z0-9_]', '', name)

    # Common phrases to shorten
    replacements = {
        'stationyardsiding': 'siding',
        'for_importexport_destination_the_toc_must_ensure_that_they_have_the_necessary_loading_offloading_capacity_for_crossboarder_traffic_proper_clearance_will_be_required': '',
        'are_you_going_to_use_the_same_types_of_wagons_for_the_entire_route_yes_or_no_if_yes_then_indicate_one_wagon_type_if_no_then_indicate_multiple_wagon_types': 'type',
        'are_you_going_to_have_the_same_train_length_for_the_entire_route_yes_or_no_if_yes_indicate_one_train_length_if_no_indicate_the_multiple_train_lengths': 'length',
        'are_you_going_to_use_the_same_loco_class_for_the_entire_route_yes_or_no_if_yes_then_indicate_one_locomotive_class_if_no_then_indicate_multiple_locomotive_classes': 'class',
        'please_provide_details_of_your_siding_operation_and_handling_methodology_per_siding__please_list_any_constraints_on_siding_operations_for_specific_times_of_the_day_and_week_eg_no_lighting_at_night_or_halfshift_working_on_weekends_etc_maximum_sizes_of_wagon_batches_that_can_be_accepted_and_handled_per_time_period_loading_and_offloading_equipment_to_be_used_and_minimum__maximum_loading_and_offloading_capacities_and_tempos': 'operating_hours_details',
        'please_provide_details_of_the_commodity_to_be_railed_stating_commodity_grade_commodity_handling_requirements_any_safety_risks_related_to_handling_or_railing_this_commodity_any_auxiliary_equipment_and_applicable_licenses_required_for_this_commodity__please_provide_proof_of_application_with_rsr': 'commodity_details',
        'dangerous_abnormal_loads_goods_details_include_the_relevant_information_consigner_consignee_cargo_freight_name_frequency_of_transportation_quantity_to_be_transported__teu_dispatching_and_receiving_destinations_train_drivers__hazmat_awareness_competent': 'dangerous_goods_details',
        'description_of_the_crewing_methodology_used_by_the_applicant_eg_bookoff_cross_point_working_round_trip_working_etc': 'crewing_methodology',
        'list_of_en_route_train_configuration_changes_eg_en_route_locations_where_train_length_will_change_where_train_will_be_split_or_combined_etc': 'train_config_changes',
        'rail_yard_and_capacity_required_for_all_relevant_yards_on_the_route': 'rail_yard_capacity',
        'uic_commodity_codes_for_office_use_only': 'uic_commodity_codes',
        'applicant_track_access_charge_payment_method': 'payment_method',
        'locomotive_consist_how_many_locomotives_per_type_will_run_in_a_consist_consist_composition_description': 'locomotive_consist',
        'maximum_power_usage_of_locomotive_consist_in_kw_hours_for_electrical_locomotives': 'locomotive_power_usage',
        'traction_power_per_locomotive_type_used_in_kilonewton': 'locomotive_traction_power',
        'wagon_type_are_you_going_to_use_the_same_types_of_wagons_for_the_entire_route_yes_or_no_if_yes_then_indicate_one_wagon_type_if_no_then_indicate_multiple_wagon_types': 'wagon_type',
        'wagon_length_of_each_wagon_type_used': 'wagon_length',
        'wagon_tare_of_each_wagon_type_used': 'wagon_tare',
        'wagon_payload_of_each_wagon_type_used': 'wagon_payload',
        'list_of_applicants_rolling_stock_maintenance_depots_and_their_locations': 'maintenance_depots',
        'train_type_freight_passenger': 'train_type',
        'brake_type_airbrake_vacuum_brake_dual': 'brake_type',
        'commodities_to_be_transported_full_description_of_the_commodity': 'commodity_description',
        'specify_commodity_environmental_risks_please_indicate_the_annexures_submitted': 'environmental_risks',
        'locomotive_type_diesel_or_electric_or_dieselelectric': 'locomotive_type',
        'gross_locomotive_mass': 'locomotive_mass',
        'maximum_train_length_meters': 'max_train_length',
        'minimum_train_length_meters': 'min_train_length',
        'gross_train_mass_in_tons': 'gross_train_mass',
        'volumes_forecast': 'volumes_forecast',
        'ancillary_services_required': 'ancillary_services',
        'slot_request_period_starting_date_yyyymmdd': 'slot_start_date',
        'slot_request_period_completion_date_yyyymmdd': 'slot_end_date',
        'frequency_required_for_the_forward_leg': 'forward_leg_frequency',
        'frequency_required_for_the_return_leg': 'return_leg_frequency',
        'number_of_wagons_per_train': 'wagons_per_train',
        'route_origin_siding_number': 'origin_siding_number',
        'route_destination_siding_number': 'destination_siding_number',
        'route_origin_siding': 'origin_siding',
        'route_destination_siding': 'destination_siding',
        'applicant_operating_hours_hhmm': 'operating_hours',
        'payment_method': 'payment_method_details',
        '1': 'payment_method_1',
        '2': 'payment_method_2',
        '3': 'payment_method_3',
        'vryheid': 'vryheid',
        'ermelo': 'ermelo',
        'location': 'location',
        'commodities_to_be_transported': 'commodities',
        'commodity_type': 'commodity_type',
        'train_details': 'train_details',
        'wagon_details': 'wagon_details',
        'locomotive_details': 'locomotive_details',
        'traction_type': 'traction_type',
        'applicant_name': 'applicant_name',
        'applicant_website': 'applicant_website',
        'company_registration_number': 'company_registration_number',
        'building_number': 'building_number',
        'building_name': 'building_name',
        'street': 'street',
        'city': 'city',
        'surname': 'surname',
        'name': 'name',
        'contact_number': 'contact_number',
        'email_address': 'email_address',
        'source_file': 'source_file'
    }

    for old, new in replacements.items():
        name = name.replace(old, new)
    
    # Remove any remaining double underscores or trailing/leading underscores
    name = name.replace('__', '_').strip('_')

    return name

def extract_data_from_docx(file_path):
    """Extracts key-value pairs from tables in a .docx file."""
    data = {"source_file": os.path.basename(file_path)}
    try:
        document = docx.Document(file_path)
        for table in document.tables:
            for row in table.rows:
                if len(row.cells) >= 2:
                    key = clean_text(row.cells[0].text)
                    value = clean_text(row.cells[1].text)
                    if key:
                        data[key] = value
    except Exception as e:
        print(f"Error processing {file_path}: {e}")
    return data

def extract_data_from_doc(file_path):
    """Extracts key-value pairs from the text of a .doc file."""
    data = {"source_file": os.path.basename(file_path)}
    try:
        word = win32.gencache.EnsureDispatch('Word.Application')
        doc = word.Documents.Open(file_path, ReadOnly=True)
        text = doc.Content.Text
        doc.Close(constants.wdDoNotSaveChanges)
        word.Quit()

        for match in re.finditer(r"(.*?):\s*(.*?)(?=\r|\n|$)", text):
            key = clean_text(match.group(1))
            value = clean_text(match.group(2))
            if key and value and len(key) < 100: # Avoid overly long keys
                data[key] = value
    except Exception as e:
        print(f"Error processing {file_path}: {e}")
    return data

def main():
    """
    Recursively reads Word application forms, extracts cleaned key-value data,
    and combines it into a single pandas DataFrame.
    """
    data_dir = r"C:\Users\ereit\OneDrive - Transnet SOC Ltd\Documents\_2025\19 Import application forms\Coal BU"
    all_forms_data = []

    # Define a list of common header/section names to exclude as columns
    header_columns_to_exclude = [
        'wagon_details',
        'train_details',
        'locomotive_details',
        'commodities_to_be_transported',
        'applicant_information_and_services_required',
        'train_configuration_and_operating_specifications',
        'payment_method',
        'uic_commodity_codes',
        'vryheid',
        'ermelo',
        'location',
        '1',
        '2',
        '3'
    ]

    for root, _, files in os.walk(data_dir):
        for file in files:
            file_path = os.path.join(root, file)
            if file.lower().endswith(".docx"):
                form_data = extract_data_from_docx(file_path)
                if len(form_data) > 1:
                    all_forms_data.append(form_data)
            elif file.lower().endswith(".doc"):
                form_data = extract_data_from_doc(file_path)
                if len(form_data) > 1:
                    all_forms_data.append(form_data)

    if all_forms_data:
        df = pd.DataFrame(all_forms_data)

        # --- Data Cleaning and Preprocessing ---
        # Apply clean_text to all cell values
        df = df.map(clean_text) # Changed from applymap to map

        # 1. Clean column names and apply shortening
        df.columns = [clean_column_name(col) for col in df.columns]

        # 2. Drop columns that are mostly NaN or are identified as headers
        # Drop if all values are NaN
        df.dropna(axis='columns', how='all', inplace=True)
        # Drop if more than 50% of values are NaN
        df.dropna(axis='columns', thresh=int(0.5 * len(df)), inplace=True)

        # Drop columns identified as headers
        df.drop(columns=[col for col in df.columns if col in header_columns_to_exclude], errors='ignore', inplace=True)

        # 3. Consolidate duplicate columns (after cleaning names)
        # Create a new DataFrame to store consolidated data
        consolidated_df = pd.DataFrame()
        processed_cols = set()

        for col in df.columns:
            if col not in processed_cols:
                # Find all columns that map to the same cleaned name
                # Use the cleaned column name for comparison
                matching_cols = [c for c in df.columns if clean_column_name(c) == clean_column_name(col)]
                
                if matching_cols:
                    # Take the first non-null value across matching columns
                    # Use .loc to avoid SettingWithCopyWarning
                    consolidated_df.loc[:, col] = df[matching_cols].bfill(axis=1).iloc[:, 0]
                    for mc in matching_cols:
                        processed_cols.add(mc)
                else:
                    # If no matching columns (shouldn't happen if all columns are processed), just add it
                    consolidated_df.loc[:, col] = df[col]
                    processed_cols.add(col)
        
        df = consolidated_df

        print("Successfully extracted and processed data from application forms.")
        print(f"Found {len(df)} forms.")
        print("Cleaned columns:", df.columns.tolist())
        print("\nSample of cleaned data:")
        print(df.head().to_string())
        
        # df.to_csv("final_extracted_data.csv", index=False)
    else:
        print("No data could be extracted from the forms.")

if __name__ == "__main__":
    main()