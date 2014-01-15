__author__ = 'Janos'

import sys
from excel_model_evaluator import ExcelDriverCOM, ExcelBlackBoxEvaluator
import csv

"""
Sex	M (for males) or F (for females)	M	M or F
Age	years	45	20-79
Race	AA (for African Americans) or WH (for whites or others)	WH	AA or WH
Total Cholesterol	mg/dL	140	130-320	170
HDL-Cholesterol	mg/dL	50	20-100	50
Systolic Blood Pressure	mm Hg	115	90-200	110
Treatment for High Blood Pressure	Y (for yes) or N (for no)	N	Y or N	N
Diabetes	Y (for yes) or N (for no)	N	Y or N	N
Smoker	Y (for yes) or N (for no)	N	Y or N	N
"""

input_dict = {"Sex": ("C3","Omnibus"),
                      "Age years": ("C4", "Omnibus"),
                      "Race": ("C5", "Omnibus"),
                      "Total Cholesterol": ("C6", "Omnibus"),
                      "HDL-Cholesterol": ("C7", "Omnibus"),
                      "Systolic Blood Pressure": ("C8", "Omnibus"),
                      "Treatment of High Blood Pressure": ("C9", "Omnibus"),
                      "Diabetes": ("C10", "Omnibus"),
                      "Smoker": ("C11", "Omnibus")
}

output_dict = {"10-Year ASCVD Risk (%)" : ("B13", "Omnibus"),
               "Optimal 10-Year ASCVD Risk (%) - same age": ("B14", "Omnibus"),
               "Lifetime ASCVD Risk (%)": ("B16", "Omnibus"),
               "Optimal Lifetime ASCVD Risk (%) - same age": ("B17", "Omnibus")
}

ranges_to_evaluate = {"Sex": ("M", "F"),
                      "Age years": (45,50,55,60,65),
                      "Race": ("AA","WH"),
                      "Total Cholesterol": (130, 150, 180, 200, 225, 320),
                      "HDL-Cholesterol": (40, 60, 80, 100),
                      "Systolic Blood Pressure": (90, 110, 130, 150, 170, 190),
                      "Treatment of High Blood Pressure": ("Y", "N"),
                      "Diabetes": ("Y", "N"),
                      "Smoker": ("Y", "N")
                      }

order_to_evaluate = ["Sex", "Age years", "Race", "Total Cholesterol", "HDL-Cholesterol", "Systolic Blood Pressure",
                     "Treatment of High Blood Pressure", "Diabetes", "Smoker"]

output_order_to_read = ["10-Year ASCVD Risk (%)", "Optimal 10-Year ASCVD Risk (%) - same age",
                        "Lifetime ASCVD Risk (%)", "Optimal Lifetime ASCVD Risk (%) - same age"]




def main(input_spreadsheet_path, csv_output_file_name):

    total_evals_i = 1
    for variable in order_to_evaluate:
        total_evals_i *= len(ranges_to_evaluate[variable])

    print("Will make %s evaluations of the spreadsheet" % total_evals_i)



    i = 0
    sex = ranges_to_evaluate["Sex"]
    age_years = ranges_to_evaluate["Age years"]
    race = ranges_to_evaluate["Race"]
    tc = ranges_to_evaluate["Total Cholesterol"]
    hdl = ranges_to_evaluate["HDL-Cholesterol"]
    sbp = ranges_to_evaluate["Systolic Blood Pressure"]
    bpt = ranges_to_evaluate["Treatment of High Blood Pressure"]
    diabetes = ranges_to_evaluate["Diabetes"]
    smoker = ranges_to_evaluate["Smoker"]



    with open(csv_output_file_name, "wb") as fw:
        csv_writer = csv.writer(fw)
        csv_writer.writerow(["row"] + order_to_evaluate + output_order_to_read)
        for j in range(len(sex)):
            for k in range(len(age_years)):
                for l in range(len(race)):
                    for m in range(len(tc)):
                        for n in range(len(hdl)):
                            for o in range(len(sbp)):
                                for p in range(len(bpt)):
                                    for q in range(len(diabetes)):
                                        for r in range(len(smoker)):

                                            dict_to_evaluate = {"Sex": sex[j], "Age years": age_years[k],
                                                                "Race": race[l], "Total Cholesterol": tc[m],
                                                                "HDL-Cholesterol": hdl[n],
                                                                "Systolic Blood Pressure": sbp[o],
                                                                "Treatment of High Blood Pressure": bpt[p],
                                                                "Diabetes": diabetes[q],
                                                                "Smoker": smoker[r],
                                                                }

                                            row_to_write = []
                                            for variable in order_to_evaluate:
                                                row_to_write += [dict_to_evaluate[variable]]

                                            if i % 100 == 0:
                                                if i == 0:
                                                    pass
                                                else:
                                                    excel_black_obj.close()

                                                excel_black_obj = ExcelBlackBoxEvaluator(input_dict, output_dict, input_spreadsheet_path, ExcelDriverCOM)

                                            result_dict = excel_black_obj.evaluate(dict_to_evaluate)

                                            result_list = []
                                            for output_variable in output_order_to_read:
                                                result_list += [result_dict[output_variable]]

                                            csv_writer.writerow([i] + row_to_write + result_list)
                                            if i == 201:
                                                excel_black_obj.close()
                                                exit()

                                            i += 1


if __name__ == "__main__":
    if len(sys.argv) == 3:
        main(sys.argv[1],sys.argv[2])
    else:
        main("C:\\Users\\Administrator\\Desktop\\Copy of Omnibus Spreadsheet NEW.xls","C:\\temp\\cardiac_risk_eval.csv")


