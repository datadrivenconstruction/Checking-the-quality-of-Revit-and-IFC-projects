import time
from datetime import date
from fpdf import FPDF
import os
import re
import pandas as pd
from .converter import convert
import matplotlib
import matplotlib.pyplot as plt
import missingno as msno
import traceback
matplotlib.use('Agg')


class Report:
    def __init__(
                    self,
                    queue,
                    data
            ):

        try:
            self.run(queue, data)
        except Exception as e:
            queue.put(traceback.format_exc())
            time.sleep(3)
        finally:
            queue.put('done')

    def run(self, queue, data):

        from .capture import ImgCapture
        ImgCapture = ImgCapture()

        queue = queue



        pathToFolder = data["pathToFolder"]
        pathPdfSources = data["pathPdfSources"]
        varGroupingParam = data["varGroupingParam"]
        pathFileExcel = data["pathFileExcel"]
        pathOutputFolder = data["pathOutputFolder"]
        pathConverterFolder = data["pathConverterFolder"]
        varProjectType = data["varProjectType"]
        varCheckSubfolders = data["varCheckSubfolders"]

        varProjectType = {"Revit": '.rvt', 'IFC': '.ifc'}[varProjectType]

        today = date.today()
        d2 = today.strftime("%d.%m.%Y")

        df_param = pd.read_excel(pathFileExcel)

        if varCheckSubfolders == "Yes":
            files = [{"name": name, "path": os.path.join(root, name)} for root, dirs, files_list in
                     os.walk(pathToFolder) for name in files_list if name.endswith(varProjectType)]
        else:
            files = [{"name": file, "path": os.path.join(pathToFolder, file)} for file in os.listdir(pathToFolder) if file.endswith(varProjectType)]

        csv_files = convert(queue, files, pathOutputFolder, pathConverterFolder, varProjectType)

        if varProjectType == ".rvt":
            listpar = list(zip(df_param['REVIT category'], df_param['Parameters to check']))
        else:
            listpar = list(zip(df_param['IFC category'], df_param['Parameters to check']))

        for file in csv_files:
            pathf = os.path.join(pathOutputFolder, file[:-4] + '_IMG')

            try: os.mkdir(pathf)
            except: pass

            file_dae = os.path.join(pathOutputFolder, file.replace('.csv', '.dae'))
            try:
                ImgCapture.load(file_dae)
            except:
                queue.put(traceback.format_exc())

            dictl = {}
            for el in listpar:
                dictl.setdefault(el[0], []).append(el[1])

            class PDF(FPDF):
                def header(self):
                    self.set_font('Arial', 'B', 12)
                    self.set_text_color(56, 81, 153)
                    self.cell(1, 0, '', 50, 100, 'L')
                    self.cell(1, 6, 'BIM PARAMETER CHECK', 50, 1, 'L', link='')
                    self.cell(1, 6, projectf[:-4], 50, 200, 'L')
                    self.image(os.path.join(pathPdfSources, 'logo_odb.png'), 25, 3, 175)
                    self.ln(6)
                    self.set_text_color(0, 0, 0)

                def footer(self):
                    self.set_y(-20)
                    self.set_font('Arial', '', 8)
                    self.image(os.path.join(pathPdfSources, 'footer_odb.png'), x=24, y=None, w=170, h=0, type='',
                               link='https://opendatabim.io/')
                    self.set_text_color(56, 81, 153)
                    self.cell(170, 4, d2 + '  ' + 'Page  ' + str(self.page_no()) + '/{nb}', 0, 0, 'R')

            pdf = PDF()
            pdf.alias_nb_pages()

            pdf.set_left_margin(25)
            pdf.set_right_margin(15)

            # Project name in the document header
            projectf = file
            strnr = 1

            if varProjectType == ".rvt":
                df = pd.read_csv(os.path.join(pathOutputFolder, file), low_memory=False)
            else:
                try: df = pd.read_csv(os.path.join(pathOutputFolder, file), low_memory=False)
                except: df = pd.read_csv(os.path.join(pathOutputFolder, file), low_memory=False, encoding="ISO-8859-1")

            ent_value = df[varGroupingParam].unique()
            ent_value = [item for item in ent_value if not pd.isnull(item)]

            klnummer = 1

            # Going through all categories of the project
            for elkl in ent_value:
                if elkl in dictl.keys():

                    # Creating a data frame for a specific group
                    klasse_dfmatch, klasse_dfmatchw = [], []
                    dfparam = df.loc[df[varGroupingParam] == elkl]
                    dfparam = dfparam.dropna(axis=1, how='all')
                    klasse_dfmatch = dfparam.columns
                    df_klasse = df.loc[:, klasse_dfmatch]
                    df_klasse = df_klasse.loc[df_klasse[varGroupingParam] == elkl]
                    df_klasse = df_klasse.dropna(how='all', axis=0)

                    # Receiving data that does not meet the criteria
                    n_inproject = []
                    odb_param = []
                    if elkl in dictl.keys():
                        parameter_tab = dictl[elkl]
                        param_len = len(parameter_tab)
                        for el in df_klasse.columns[:45].to_list():
                            if el in parameter_tab:
                                odb_param.append(el)
                            else:
                                n_inproject.append(el)

                    # Obtaining the table of parameters distributed by the number of values
                    nval = [df_klasse[col_name].count() for col_name in df_klasse.columns]
                    dict_nv = {}
                    dict_nv = dict(zip(df_klasse.columns, nval))
                    dict_nv = sorted(dict_nv.items(), key=lambda x: x[1], reverse=True)

                    # Parameters that are not in the table to check
                    dict_nvn = []
                    for el in dict_nv:
                        dict_nvn.append(el[0])
                    for el in odb_param:
                        dict_nvn.remove(el)

                    for el in odb_param:
                        dict_nvn.insert(0, el)
                    df_klasse = df_klasse[dict_nvn]

                    # Parameters found in XLSX table to check
                    parameter_s23 = parameter_tab
                    for el23 in odb_param:
                        parameter_s23.remove(el23)

                    # Defining colors for parameters that are found and parameters that were not found
                    colorsd = []
                    for el in odb_param:
                        colorsd.append('#43CD62')
                    for el in range(len(dict_nvn)):
                        colorsd.append('grey')
                    # Creating a picture PNG for all parameters
                    namefig = msno.bar(df_klasse[df_klasse.columns[:45]], figsize=(25, 4), fontsize=20,
                                       color=colorsd)
                    fig_copy = namefig.get_figure()
                    elklclean = re.sub('[^A-Za-z0-9]+', '', elkl)

                    # Save the picture in the folder
                    fig_copy.savefig(os.path.join(pathf, elklclean + '_attributesklasse.png'), bbox_inches='tight')
                    namefig.clear()
                    plt.clf()

                    df_klasse = df_klasse.set_index('ID')
                    df_klasse.index = df_klasse.index.astype(str)
                    nodes = list(df_klasse.index)

                    # Forming geometry for elements in the group at the top of the document
                    level_img = os.path.join(pathOutputFolder, file[:-4] + '_IMG',
                                             file[:30] + elklclean + '_geometry_image.png')
                    img = ImgCapture.shoot(level_img, nodes)

                    # bookm[klne + ' ' + varGroupingParam + ': '+ elkl]=pdf.page_no()

                    ################################################################
                    ###################### PDF document ############################
                    ################################################################

                    # Creating a new page of a PDF document
                    pdf.add_page()
                    klnummer += 1

                    ################################################################
                    # If the group photo is generated
                    try:
                        pdf.image(level_img, x=55, y=32, w=65, type='')
                    except:
                        pass
                    pdf.set_font('Arial', 'B', 10)
                    pdf.set_text_color(0, 102, 204)
                    pdf.set_font('', 'U')
                    pdf.set_font('Arial', '', 10)
                    pdf.set_text_color(0, 0, 0)

                    ################################################################

                    # Group name at the top of the document
                    klne = str(strnr) + '.' + str(klnummer)
                    pdf.set_font('Arial', 'B', 12)
                    pdf.set_fill_color(241, 241, 242)
                    pdf.cell(len(varGroupingParam) + 24, 7, ' ' + klne + ' ' + varGroupingParam + ': ', 0, 0, 'L', fill=True)
                    pdf.cell(3, 7, ' ', 0, 0, 'L')
                    pdf.set_fill_color(210, 255, 224)
                    pdf.cell(len(elkl) + 20, 7, ' ' + elkl, 0, 1, 'C', fill=True)
                    pdf.ln(57)

                    ################################################################

                    pdf.set_font('Arial', '', 10)
                    # Check the number of parameters found, if more than 50%
                    if len(odb_param) / param_len > 0.5:
                        pdf.image(os.path.join(pathPdfSources, 'v.png'), x=125, y=57, w=75)
                    else:
                        pdf.image(os.path.join(pathPdfSources, 'w.png'), x=130, y=58, w=70)

                    ################################################################

                    # Forming a table with the statistics of the parameters checked and as a whole

                    pdf.set_font('Arial', '', 8)
                    pdf.cell(32, 7, 'Checking parameters by criteria from an Excel table:', 0, 1, 'L')

                    pdf.set_font('Arial', 'B', 10)
                    pdf.set_fill_color(225, 225, 225)
                    pdf.cell(32, 7, 'Elements', 0, 0, 'C', fill=True)
                    pdf.cell(3, 7, '', 0, 0, 'c')
                    pdf.cell(32, 7, 'Parameters', 0, 0, 'C', fill=True)
                    pdf.cell(3, 7, '', 0, 0, 'c')
                    pdf.cell(32, 7, 'Checked', 0, 0, 'C', fill=True)
                    pdf.cell(3, 7, '', 0, 0, 'c')
                    pdf.set_fill_color(165, 250, 184)
                    pdf.cell(32, 7, 'Project has', 0, 0, 'C', fill=True)
                    if len(parameter_s23) == 0:
                        pdf.set_fill_color(165, 255, 200)
                    else:
                        pdf.set_fill_color(250, 180, 165)
                    strparam = str(parameter_s23)
                    pdf.cell(3, 7, '', 0, 0, 'c')
                    strparam = strparam.replace("[", "")
                    strparam = strparam.replace("]", "")
                    pdf.cell(32, 7, 'Missing', 0, 0, 'C', fill=True)
                    pdf.set_font('Arial', '', 11)
                    pdf.cell(20, 8, '', 2, 1, 'L')

                    ##############

                    pdf.set_font('Arial', '', 10)
                    pdf.set_fill_color(241, 241, 242)
                    pdf.cell(32, 7, str(len(df_klasse)), 0, 0, 'C', fill=True)
                    pdf.cell(3, 7, '', 0, 0, 'c')
                    pdf.cell(32, 7, str(len(df_klasse.columns)), 0, 0, 'C', fill=True)
                    pdf.cell(3, 7, '', 0, 0, 'c')
                    pdf.cell(32, 7, str(param_len), 0, 0, 'C', fill=True)
                    pdf.cell(3, 7, '', 0, 0, 'c')
                    pdf.set_fill_color(210, 255, 224)
                    pdf.cell(32, 7, str(len(odb_param)), 0, 0, 'C', fill=True)
                    if len(parameter_s23) == 0:
                        pdf.set_fill_color(165, 255, 200)
                    else:
                        pdf.set_fill_color(255, 218, 210)
                    pdf.cell(3, 7, '', 0, 0, 'c')
                    pdf.cell(32, 7, str(len(parameter_s23)), 0, 0, 'C', fill=True)
                    pdf.set_font('Arial', '', 11)
                    pdf.cell(20, 10, '', 2, 1, 'L')

                    ################################################################

                    # Parameters not found in the project
                    if len(parameter_s23) > 0:
                        pdf.cell(150, 9, 'Parameters not found in the project or completely empty values:', 0, 1,
                                 'L')
                        pdf.set_font('Arial', '', 10)
                        for el in parameter_s23:
                            pdf.set_fill_color(255, 218, 210)
                            size = len(el) + 20
                            pdf.cell(size, 6, el, 0, 0, 'C', fill=True)
                            pdf.cell(3, 7, '', 0, 0, 'C', )
                    else:
                        pdf.cell(150, 12, '', 0, 1, 'L')

                    ################################################################

                    # Forming a table with statistics on the volume of groups
                    pdf.ln(4)
                    pdf.set_font('Arial', '', 12)
                    try:
                        df_klasse['Volume'] = df_klasse['Volume'].astype(str).str.extract(
                            r'([-+]?\d*\.?\d+)').astype(
                            'float')
                        total_vol = int(df_klasse['Volume'].sum())
                    except:
                        total_vol = '0'
                        pass
                    try:
                        df_klasse['Area'] = df_klasse['Area'].astype(str).str.extract(r'([-+]?\d*\.?\d+)').astype(
                            'float')
                        total_area = int(df_klasse['Area'].sum())
                    except:
                        total_area = '0'
                        pass
                    try:
                        df_klasse['Length'] = df_klasse['Length'].astype(str).str.extract(
                            r'([-+]?\d*\.?\d+)').astype(
                            'float')
                        total_len = int(df_klasse['Length'].sum())
                    except:
                        total_len = '0'
                        pass
                    pdf.cell(20, 3, '', 2, 1, 'L')
                    pdf.set_font('Arial', '', 8)
                    pdf.cell(150, 7, 'Volumetric parameters for all elements in the groups:', 0, 1, 'С')

                    ########

                    pdf.set_font('Arial', 'B', 10)
                    pdf.set_fill_color(225, 225, 225)
                    pdf.cell(55, 7, 'Sum of volumes', 0, 0, 'C', fill=True)
                    pdf.cell(3, 7, '', 0, 0, 'c')
                    pdf.cell(55, 7, 'Sum of areas', 0, 0, 'C', fill=True)
                    pdf.cell(3, 7, '', 0, 0, 'c')
                    pdf.cell(55, 7, 'Sum of lengths', 0, 0, 'C', fill=True)
                    pdf.cell(3, 7, '', 0, 1, 'c')

                    ########

                    pdf.cell(20, 2, '', 2, 1, 'L')
                    pdf.set_font('Arial', '', 10)
                    pdf.set_fill_color(241, 241, 242)
                    pdf.cell(55, 7, str(total_vol) + ' m³', 0, 0, 'C', fill=True)
                    pdf.cell(3, 7, '', 0, 0, 'c')
                    pdf.cell(55, 7, str(total_area) + ' m²', 0, 0, 'C', fill=True)
                    pdf.cell(3, 7, '', 0, 0, 'c')
                    pdf.cell(55, 7, str(total_len) + ' m', 0, 0, 'C', fill=True)
                    pdf.cell(3, 7, '', 0, 1, 'c')

                    ################################################################

                    # Inserting a picture with the parameters that were checked above
                    pdf.set_font('Arial', '', 11)
                    pdf.cell(20, 5, '', 2, 1, 'L')
                    pdf.set_font('Arial', '', 11)
                    text1 = 'The following diagram shows which parameters in the Project ' + file[
                                                                                             :-9] + ' are included and which parameters and how often for ' + varGroupingParam + ' ' + elkl + ' are given:'
                    pdf.multi_cell(w=160, h=6, txt=text1, border=0)
                    pdf.cell(20, 1, '', 2, 1, 'L')
                    pdf.set_font('Arial', 'B', 12)
                    pdf.image(os.path.join(pathPdfSources, 'logo_atribbuten.png'), x=30, w=165, type='', link='')
                    pdf.image(os.path.join(pathf, elklclean + '_attributesklasse.png'), x=25, h=50, w=165, type='',
                              link='')

                    ################################################################

                    # Formation of the bottom table of statistics for the found parameters
                    # Calculate the number of non-null values for each column in the dataframe using the
                    non_null_counts = dfparam.count()
                    # Calculate the total number of rows in the dataframe using "len()", and store the result in "total_counts"
                    total_counts = len(dfparam)
                    percent_filled = non_null_counts / total_counts * 100
                    # Reset the index of the dataframe "filled_df"
                    filled_df = pd.DataFrame(percent_filled).reset_index()
                    filled_df.columns = ['column_name', 'Percentage filling']
                    # Convert the "Percentage filling" column to integer type
                    filled_df['Percentage filling'] = filled_df['Percentage filling'].astype(int)

                    for col in dfparam.columns:
                        unique_values = dfparam[col].unique()[:5]
                        filled_df.loc[filled_df['column_name'] == col, 'Example values'] = ", ".join(
                            map(str, unique_values))

                    filled_df['Example values'] = filled_df['Example values'].str.replace('[|]|nan ,|nan,|nan', '')
                    filled_df['Example values'] = filled_df['Example values'].str[:50]
                    filled_df = filled_df.set_index('column_name')

                    filled_df2 = filled_df.loc[odb_param]
                    pdf.ln(4)

                    ###############
                    if len(filled_df2) > 0:
                        pdf.set_font('Arial', '', 8)
                        pdf.cell(150, 7, 'Statistics on the parameters found:', 0, 1, 'С')
                        pdf.set_font('Arial', 'B', 10)
                        pdf.set_fill_color(165, 250, 184)
                        pdf.cell(55, 6, txt='  Parameter Name', ln=0, align="L", fill=True)
                        pdf.cell(3, 6, '', ln=0, align="C")
                        pdf.set_fill_color(225, 225, 225)
                        pdf.cell(35, 6, txt='Percentage filling', ln=0, align="C", fill=True)
                        pdf.cell(3, 6, '', ln=0, align="C")
                        pdf.cell(70, 6, txt='  Example values', ln=0, align="L", fill=True)
                        pdf.cell(3, 7, '', ln=1, align="C")

                        for i in range(len(filled_df2)):
                            pdf.set_font('Arial', '', 10)
                            pdf.set_fill_color(210, 255, 224)
                            pdf.cell(55, 6, txt='  ' + str(filled_df2.index[i][:38]), ln=0, align="L", fill=True)
                            if filled_df2.iloc[i, 0] < 50:
                                pdf.set_fill_color(255, 218, 210)
                            else:
                                pdf.set_fill_color(210, 255, 224)
                            pdf.cell(3, 6, '', ln=0, align="C")
                            pdf.cell(35, 6, txt=str(filled_df2.iloc[i, 0]) + '%', ln=0, align="C", fill=True)
                            pdf.set_fill_color(241, 241, 242)
                            pdf.cell(3, 6, '', ln=0, align="C")
                            pdf.cell(70, 6, txt='  ' + str(filled_df2.iloc[i, 1][:38]), ln=0, align="L", fill=True)
                            pdf.cell(3, 7, '', ln=1, align="C")

            strnr += 1
            pdf.output(os.path.join(pathOutputFolder, file[:-4] + '_ODB_Report.pdf'), 'F')
            queue.put('Report generated: ' + file[:-4] + '_ODB_Report.pdf')
            time.sleep(1)
