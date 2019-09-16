import os
import pandas as pd

# read data
data = pd.read_excel('sample.xlsx')
for i in range(len(data)):
    name = data['نام و نام خانوادگی'].iloc[i]
    print(name)
    code = data['کد'].iloc[i]
    print(code)
    FBS = data['FBS'].iloc[i]
    Urea = data['Urea'].iloc[i]
    creat = data['creat'].iloc[i]
    UA = data['UA'].iloc[i]
    TG = data['TG'].iloc[i]
    Chol = data['Chol'].iloc[i]
    HDL_C = data['HDL-C'].iloc[i]
    LDL_C = data['LDL-C'].iloc[i]
    Ca = data['Ca'].iloc[i]
    AST = data['AST'].iloc[i]
    ALT = data['ALT'].iloc[i]
    ALK_P = data['ALK.P'].iloc[i]
    B_D = data['B.D'].iloc[i]
    B_T = data['B.T'].iloc[i]

    file_name = str(code)+'.tex'
    with open(file_name,'w',encoding='utf8') as file:
        file.write('\\documentclass[a4paper , 12pt]{article}\n')
        file.write('\\usepackage{geometry} \n')
        file.write('\\usepackage{fancyhdr} \n')
        file.write('\\usepackage{graphicx} \n')
        file.write('\\usepackage[table,xcdraw]{xcolor} \n')
        file.write('\\usepackage{xepersian} \n')
        file.write('\\pagestyle{fancy}')
        file.write('\\setlength\\headheight{30 pt}')
        file.write('\\rhead{\\includegraphics[scale = 0.02]{iran}} \n')
        file.write('\\lhead{\\includegraphics[scale = 0.02]{iran}} \n')
        file.write('\\renewcommand{\\headrulewidth}{1pt} \n')
        file.write('\\renewcommand{\\footrulewidth}{1pt} \n')
        file.write('\\setlatintextfont{Times New Roman}\n')
        file.write('\\setlatintextfont{Arial} \n')
        file.write('\\settextfont{XB Niloofar} \n')
        file.write('\\begin{document} \n')
        file.write('\\hspace{2cm} \n')
        file.write('\\begin{center} \\section*{نتایج آزمایشات } \\end{center} \n')
        file.write('\\textbf{نام و نام خانوادگی :}' + '{}'.format(name) + '\\hspace{4cm}' +'\\textbf{کد ثبت اطلاعات :}' + '{}'.format(code) + '\n')
        file.write('\\begin{latin} \n')
        file.write('\\begin{table}[h] \n')
        file.write('\\centering \n')
        file.write('\\begin{tabular}{|c|c|c|c|} \n')
        file.write('\\hline \n')
        file.write('\\textbf{Test} & \\textbf{Result} & \\textbf{Units} & \\textbf{Reference Value} \\\\ \\hline \n')
        if 70 <= FBS <= 110:
            file.write('\\begin{tabular}[c]{@{}c@{}}Fasting Blood Sugar\\\\ (FBS)\\end{tabular} & \\multicolumn{1}{c|}{'+ str(FBS) +'} & \\multicolumn{1}{l|}{Mg/dl} & 70 - 110 \\\\ \\hline \n')
        else:
            file.write('\\begin{tabular}[c]{@{}c@{}}Fasting Blood Sugar\\\\ (FBS)\\end{tabular} & \\multicolumn{1}{c|}{\cellcolor{gray!50}'+ str(FBS) +'} & \\multicolumn{1}{l|}{Mg/dl} & 70 - 110 \\\\ \\hline \n')
        if 10 <= Urea <= 45:
            file.write('Urea & '+ str(Urea) +'& Mg/dl & 10 - 45 \\\\ \\hline \n')
        else:
            file.write('Urea & \cellcolor{gray!50}' + str(Urea) + '& Mg/dl & 10 - 45 \\\\ \\hline \n')
        if 0.5 <= creat <= 1.4:
            file.write('Creatinine & '+ str(creat)+'  & Mg/dl & 0.5 - 1.4 \\\\ \\hline \n')
        else:
            file.write('Creatinine & \cellcolor{gray!50}' + str(creat) + '  & Mg/dl & 0.5 - 1.4 \\\\ \\hline \n')
        file.write('Uric Acid &'+ str(UA) +'  & Mg/dl & \\begin{tabular}[c]{@{}c@{}}Females :2.3 - 6.1 \\\\ Males : 3.6 - 8.2 \\end{tabular} \\\\ \\hline \n')
        if Chol <= 200:
                file.write('Cholesterol & '+ str(Chol) + ' & Mg/dl & Up to 200 \\\\ \\hline \n')
        else:
            file.write('Cholesterol & \cellcolor{gray!50}' + str(Chol) + ' & Mg/dl & Up to 200 \\\\ \\hline \n')
        if TG <= 200:
            file.write('Triglyceride &  ' + str(TG) + ' & Mg/dl & Up to 200 \\\\ \\hline \n')
        else:
            file.write('Triglyceride & \cellcolor{gray!50} ' + str(TG) + ' & Mg/dl & Up to 200 \\\\ \\hline \n')
        if HDL_C >= 35:
            file.write('HDL - Cholesterol & '+ str(HDL_C) + ' & Mg/dl & > 35 \\\\ \\hline \n')
        else:
            file.write('HDL - Cholesterol & \cellcolor{gray!50}' + str(HDL_C) + ' & Mg/dl & > 35 \\\\ \\hline \n')
        if LDL_C <= 150:
            file.write('LDL - Cholesterol & '+ str(LDL_C) + ' & Mg/dl & < 150 \\\\ \\hline \n')
        else:
            file.write('LDL - Cholesterol & \cellcolor{gray!50}'+ str(LDL_C) + ' & Mg/dl & < 150 \\\\ \\hline \n')
        file.write('AST & '+ str(AST) + ' & U/L & \\begin{tabular}[c]{@{}c@{}}Females : < 31 \\\\ Males : < 37 \\end{tabular} \\\\ \\hline \n')
        file.write('ALT & '+ str(ALT) + ' & U/L & \\begin{tabular}[c]{@{}c@{}}Females : < 31\\\\ Males : < 41 \\end{tabular} \\\\ \\hline \n')
        file.write('ALP & '+ str(ALK_P) + ' & U/L & \\begin{tabular}[c]{@{}c@{}}Adult Females : 64 - 306\\\\ Adult Males : 80 - 306\\end{tabular} \\\\ \\hline \n')
        if B_D <= 0.6:
            file.write('Bilirubin Direct & '+ str(B_D) + ' & Mg/dl & Up to 0.6 \\\\ \\hline \n')
        else:
            file.write('Bilirubin Direct & \cellcolor{gray!50}'+ str(B_D) + ' & Mg/dl & Up to 0.6 \\\\ \\hline \n')
        if B_T <= 1.1:
            file.write('Bilirubin Total & '+ str(B_T) + ' & Mg/dl & Up to 1.1 \\\\ \\hline \n')
        else:
            file.write('Bilirubin Total &  \cellcolor{gray!50}' + str(B_T) + ' & Mg/dl & Up to 1.1 \\\\ \\hline \n')
        if 8.1 <= Ca<= 10.4:
            file.write('Calcium & '+ str(Ca) + ' & Mg/dl & 8.1 - 10.4 \\\\ \\hline \n')
        else:
            file.write('Calcium & \cellcolor{gray!50}' + str(Ca) + ' & Mg/dl & 8.1 - 10.4 \\\\ \\hline \n')
        file.write('Vitamin D & '+ str('-') + ' & Mg/ml & \\begin{tabular}[c]{@{}c@{}}Deficiency : < 10\\\\ insufficiency : 10 - 30\\\\ sufficiency : 30 - 100\\\\ toxicity : > 100\\end{tabular} \\\\ \\hline \n')
        file.write('\\end{tabular} \n')
        file.write('\\end{table} \n')
        file.write('\\end{latin} \n')
        file.write('\\end{document}')
        file.close()
    commend = 'xelatex {}'.format(file_name)
    os.system(commend)

