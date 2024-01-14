import pandas as pd
import numpy as np
import os
class Analysis:
    def __init__(self, 
                 well_number, 
                 total_count, 
                 phl_count, 
                 yemk_count, 
                 live_percentage, 
                 dead_percentage, 
                 phl_vl2, 
                 phl_bl1, 
                 yemk_vl2, 
                 yemk_bl1):
        #prior columns
        self.well_number = well_number
        self.total_count = total_count
        self.phl_count = phl_count
        self.yemk_count = yemk_count
        self.live_percentage = live_percentage
        self.dead_percentage = dead_percentage
        self.phl_vl2 = phl_vl2
        self.phl_bl1 = phl_bl1
        self.yemk_vl2 = yemk_vl2
        self.yemk_bl1 = yemk_bl1
        #created columns
        self.pHL_VL2_BL1 = self.calculate_pHL_VL2_BL1()
        self.yemk_vl2_bl1 = self.calculate_YEML_VL2_dividedby_BL1()
        #these columns are populated not at the point of creation but later
        #because they require all columns to have been created before they can be populated
        self.relative_well_number = 0
        self.slope_corrected_phl_vl2_bl1 = 0
        self.slope_corrected_yemk_vl2_bl1 = 0
        self.cutoff_PHL_VL2_BL1_below_cuttoff = ""
        self.cutoff_yemk_vl2_bl1_below_cuttoff = ""
        self.phl_z_score = 0
        self.yemk_z_score = 0
        self.live_z_score = 0
        self.hits_phl_z_score = ""
        self.hits_yemk_z_score = ""
        self.hits_live_z_score = ""

    def calculate_pHL_VL2_BL1(self):
            # Calculate pHL VL2/BL1 with 8 decimal points precision
            if self.phl_bl1 != 0:
                return round(self.phl_vl2 / self.phl_bl1, 8)
            else:
                return 0  # Handle the case where phl_bl1 is 0 to avoid division by zero
    
    def calculate_YEML_VL2_dividedby_BL1(self):
         # Calculate pHL VL2/BL1 with 8 decimal points precision
            if self.yemk_vl2 != 0:
                return round(self.yemk_vl2 / self.yemk_bl1, 8)
            else:
                return 0  # Handle the case where phl_bl1 is 0 to avoid division by zero
    
    def getRelative_wellNumber(self, i):
         self.relative_well_number = i

    def calculate_slope_corrected_phl_vl2_bl1(self, slope_phl_vl2_bl1):
        # Set self.slope_corrected_phl_vl2_bl1 = (self.pHL_VL2_BL1 - self.relative_well_number) * self.slope_phl_vl2_bl1
        self.slope_corrected_phl_vl2_bl1 = self.pHL_VL2_BL1 - (self.relative_well_number * slope_phl_vl2_bl1)

    def calculate_slope_corrected_yemk_vl2_bl1(self, slope_yemk_vl2_bl1):
        # Set self.slope_corrected_yemk_vl2_bl1 = (self.yemk_vl2_bl1 - self.relative_well_number) * self.slope_yemk_vl2_bl1
        self.slope_corrected_yemk_vl2_bl1 = self.yemk_vl2_bl1 - (self.relative_well_number * slope_yemk_vl2_bl1)

    def calculate_z_score(self, corrected_mean, corrected_SD, attribute_name):
        
        if attribute_name == "pHL_VL2_BL1":
            if self.cutoff_PHL_VL2_BL1_below_cuttoff != "":
                self.phl_z_score = (self.cutoff_PHL_VL2_BL1_below_cuttoff - corrected_mean)/corrected_SD
        elif attribute_name == "yemk_vl2_bl1":
            if self.cutoff_yemk_vl2_bl1_below_cuttoff != "":
                self.yemk_z_score = (self.cutoff_yemk_vl2_bl1_below_cuttoff - corrected_mean)/corrected_SD

    def calculate_cutoff(self, mean, sd, attribute_name):

        # Calculate the cutoff value 1.5 standard deviations above the mean
        cutoff = mean + (sd * 1.5)

        # perform both calculations
        # The value should only be present if it is below the 1.5 standard deviation above the mean cutoff.
        if attribute_name == "pHL_VL2_BL1":
            if np.absolute(self.slope_corrected_phl_vl2_bl1) > np.absolute(np.absolute(cutoff) - np.absolute(self.slope_corrected_phl_vl2_bl1)):
                self.cutoff_PHL_VL2_BL1_below_cuttoff = cutoff
        elif attribute_name == "yemk_vl2_bl1":
            if np.absolute(self.slope_corrected_yemk_vl2_bl1) > np.absolute(np.absolute(cutoff) - np.absolute(self.slope_corrected_yemk_vl2_bl1)):
                self.cutoff_yemk_vl2_bl1_below_cuttoff = cutoff

def combine_Samples_and_highControls():
    # Replace 'your_file.xlsx' with the actual path to your Excel file
    file_path = 'LC2-032_KCP1 pHL-YEMK DC 20231030.xlsx'

    # Read data from the "samples" sheet
    df_samples = pd.read_excel(file_path, sheet_name='Samples')

    # Exclude the last 2 rows
    df_filtered_samples = df_samples.iloc[:-2]

    # Create a list of Analysis instances from the filtered samples data
    analysis_samples_list = [
        Analysis(row['X1'], row['Count'], row['Live/Cells/Singlet.Cells/pHL.|.Count'],
                row['Live/Cells/Singlet.Cells/YEMK.|.Count'], row['Live.|.Freq..of.Total.(%)'],
                row['Dead.|.Freq..of.Total.(%)'], row['Live/Cells/Singlet.Cells/pHL.|.Median.(VL2-H.::.VL2-H)'],
                row['Live/Cells/Singlet.Cells/pHL.|.Median.(BL1-H.::.BL1-H)'],
                row['Live/Cells/Singlet.Cells/YEMK.|.Median.(VL2-H.::.VL2-H)'],
                row['Live/Cells/Singlet.Cells/YEMK.|.Median.(BL1-H.::.BL1-H)'])
        for _, row in df_filtered_samples.iterrows()
    ]

    # Read data from the "High Controls" sheet
    df_high_controls = pd.read_excel(file_path, sheet_name='High Controls')

    # Exclude empty rows and the last 2 rows
    df_filtered_high_controls = df_high_controls.dropna(how='all').iloc[:-2]

    # Create a list of Analysis instances from the filtered high controls data
    analysis_high_controls_list = [
        Analysis(row['X1'], row['Count'], row['Live/Cells/Singlet.Cells/pHL.|.Count'],
                row['Live/Cells/Singlet.Cells/YEMK.|.Count'], row['Live.|.Freq..of.Total.(%)'],
                row['Dead.|.Freq..of.Total.(%)'], row['Live/Cells/Singlet.Cells/pHL.|.Median.(VL2-H.::.VL2-H)'],
                row['Live/Cells/Singlet.Cells/pHL.|.Median.(BL1-H.::.BL1-H)'],
                row['Live/Cells/Singlet.Cells/YEMK.|.Median.(VL2-H.::.VL2-H)'],
                row['Live/Cells/Singlet.Cells/YEMK.|.Median.(BL1-H.::.BL1-H)'])
        for _, row in df_filtered_high_controls.iterrows()
    ]

    # Combine the lists from samples and high controls
    combined_analysis_list = analysis_samples_list + analysis_high_controls_list

    return combined_analysis_list, file_path

def create_relative_well_numbers(combined_analysis_list):
    # itterater used to create relative well number and to calculate slopes
    i = 0
    #creates relative well number and totals the numbers for the slopes
    for anaysis in combined_analysis_list:
        i+=1
        anaysis.getRelative_wellNumber(i)

def create_slope(combined_analysis_list, listX, listY, attribute_name):
    for analysis in combined_analysis_list:
        listX.append(analysis.relative_well_number)

        # Dynamically access the attribute using getattr
        attribute_value = getattr(analysis, attribute_name)
        listY.append(attribute_value)

    # Calculate the slope  the mean and standard deviation using numpy.polyfit
    slope, intercept = np.polyfit(listX, listY, 1)
    mean = np.mean(listY)
    sd = np.std(listY)

    return slope, mean, sd

def calculate_corrected_mean_SD(combined_analysis_list, listX, listY, attribute_name):
    for analysis in combined_analysis_list:
        if attribute_name == "pHL_VL2_BL1":
            if analysis.cutoff_PHL_VL2_BL1_below_cuttoff != "":
                listX.append(analysis.relative_well_number)

                # Dynamically access the attribute using getattr
                attribute_value = getattr(analysis, attribute_name)
                listY.append(attribute_value)
        elif attribute_name == "yemk_vl2_bl1":
            if analysis.cutoff_yemk_vl2_bl1_below_cuttoff != "":
                listX.append(analysis.relative_well_number)

                # Dynamically access the attribute using getattr
                attribute_value = getattr(analysis, attribute_name)
                listY.append(attribute_value)

    # calculate the mean and standard deviation using numpy.polyfit
    mean = np.mean(listY)
    sd = np.std(listY)
    return mean, sd

def populate_slope_mean_SD(combined_analysis_list):
    
    #pHL_VL2_BL1
    listX = []
    listY = []
    slope_pHL_VL2_BL1, mean_pHL_VL2_BL1, SD_pHL_VL2_BL1 = create_slope(combined_analysis_list, listX, listY, "pHL_VL2_BL1")
    
    #yemk_vl2_bl1
    listX = []
    listY = []
    slope_yemk_vl2_bl1, mean_yemk_vl2_bl1, SD_yemk_vl2_bl1 = create_slope(combined_analysis_list, listX, listY, "yemk_vl2_bl1")

    #calculate the corrected slope and populate the cutoff
    for anaysis in combined_analysis_list:
        anaysis.calculate_slope_corrected_phl_vl2_bl1(slope_pHL_VL2_BL1)
        anaysis.calculate_slope_corrected_yemk_vl2_bl1(slope_yemk_vl2_bl1)
        anaysis.calculate_cutoff(mean_pHL_VL2_BL1, SD_pHL_VL2_BL1, "pHL_VL2_BL1")
        anaysis.calculate_cutoff(mean_yemk_vl2_bl1, SD_yemk_vl2_bl1, "yemk_vl2_bl1")

    #calculate corrected mean and corrected SD
    corrected_mean_pHL_VL2_BL1, corrected_SD_pHL_VL2_BL1 = calculate_corrected_mean_SD(combined_analysis_list, listX, listY, "pHL_VL2_BL1")
    corrected_mean_yemk_vl2_bl1, corrected_SD_yemk_vl2_bl1 = calculate_corrected_mean_SD(combined_analysis_list, listX, listY, "yemk_vl2_bl1")

    for analysis in combined_analysis_list:
        analysis.calculate_z_score(corrected_mean_pHL_VL2_BL1, corrected_SD_pHL_VL2_BL1, "pHL_VL2_BL1")
        analysis.calculate_z_score(corrected_mean_yemk_vl2_bl1, corrected_SD_yemk_vl2_bl1, "yemk_vl2_bl1")

def populate_live_z_score(combined_analysis_list):
    live_percentage_list = []
    for analysis in combined_analysis_list:
        live_percentage_list.append(analysis.live_percentage)

    live_mean = np.mean(live_percentage_list)
    live_sd = np.std(live_percentage_list)

    for analysis in combined_analysis_list:
        if analysis.live_percentage != "":
            analysis.live_z_score = (analysis.live_percentage - live_mean)/live_sd

def populate_hits_phl_z_score(combined_analysis_list):
    for analysis in combined_analysis_list:
        if analysis.phl_z_score < -5:
            analysis.hits_phl_z_score = analysis.phl_z_score

def populate_hits_yemk_z_score(combined_analysis_list):
    for analysis in combined_analysis_list:
        if analysis.yemk_z_score < -5:
            analysis.yemk_z_score = analysis.yemk_z_score

def populate_hits_live_z_score(combined_analysis_list):
    for analysis in combined_analysis_list:
        if analysis.live_z_score < -5:
            analysis.hits_live_z_score = analysis.hits_live_z_score

def populate_hits(combined_analysis_list):
    populate_hits_phl_z_score(combined_analysis_list)
    populate_hits_yemk_z_score(combined_analysis_list)
    populate_hits_live_z_score(combined_analysis_list)

def write_excel_sheet(file_path):
    combined_analysis_list, file_path = combine_Samples_and_highControls()

    create_relative_well_numbers(combined_analysis_list)

    populate_slope_mean_SD(combined_analysis_list)

    populate_live_z_score(combined_analysis_list)

    populate_hits(combined_analysis_list)

    # Check if the "Analysis" sheet already exists and delete it
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
        if 'Analysis' in writer.sheets:
            writer.book.remove(writer.sheets['Analysis'])

        # Convert the combined list of Analysis instances to a DataFrame
        df_combined_analysis = pd.DataFrame([vars(a) for a in combined_analysis_list])

        # Write the DataFrame to a new sheet named "Analysis"
        df_combined_analysis.to_excel(writer, sheet_name='Analysis', index=False)

    export_All_Plates_YEMK_pHL_Live_excel_file(combined_analysis_list)
    export_All_Hits(combined_analysis_list)

def export_All_Hits(combined_analysis_list):
    # Define the file name
    excel_filename = 'All_Hits.xlsx'

    # Check if the file already exists
    if os.path.exists(excel_filename):
        # If it exists, delete the file
        os.remove(excel_filename)
        print(f"Deleted existing file: {excel_filename}")

    # Create a DataFrame from the combined analysis list
    if combined_analysis_list:
        phl_z_score_hits = []
        yemk_z_score_hits = []
        live_z_score_hits = []

        for analysis in combined_analysis_list:
            phl_z_score_hits.append(analysis.hits_phl_z_score)
            yemk_z_score_hits.append(analysis.hits_yemk_z_score)
            live_z_score_hits.append(analysis.hits_live_z_score)

        # Create a DataFrame with the specified columns
        All_Hits_columns = {
            'pHL Z-Score': phl_z_score_hits,
            'YEMK Z-Score': yemk_z_score_hits,
            'Live Z-Score': live_z_score_hits
        }

        df = pd.DataFrame(All_Hits_columns)

        # Create a new Excel file with the specified columns
        df.to_excel(excel_filename, index=False)
        print(f"Created new file: {excel_filename}")
    else:
        print("No data provided to create the Excel file.")

def export_All_Plates_YEMK_pHL_Live_excel_file(combined_analysis_list):
    # Define the file name
    excel_filename = 'All_Plates_YEMK_pHL_Live.xlsx'

    # Check if the file already exists
    if os.path.exists(excel_filename):
        # If it exists, delete the file
        os.remove(excel_filename)
        print(f"Deleted existing file: {excel_filename}")

    # Create a DataFrame from the combined analysis list
    if combined_analysis_list:
        well_number = []
        phl_z_score = []
        yemk_z_score = []
        live_z_score = []

        for analysis in combined_analysis_list:
            well_number.append(analysis.well_number)
            phl_z_score.append(analysis.phl_z_score)
            yemk_z_score.append(analysis.yemk_z_score)
            live_z_score.append(analysis.live_z_score)

        # Create a DataFrame with the specified columns
        All_Plates_YEMK_pHL_Live_columns = {
            'Well #': well_number,
            'pHL Z-Score': phl_z_score,
            'YEMK Z-Score': yemk_z_score,
            'Live Z-Score': live_z_score
        }

        df = pd.DataFrame(All_Plates_YEMK_pHL_Live_columns)

        # Create a new Excel file with the specified columns
        df.to_excel(excel_filename, index=False)
        print(f"Created new file: {excel_filename}")
    else:
        print("No data provided to create the Excel file.")

file_path = ""
write_excel_sheet(file_path)