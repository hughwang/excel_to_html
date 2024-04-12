import pandas as pd
import os

def main():
    html_string_start = '''
    <html>
      <head><title>Criminal Report</title>
      <meta charset="UTF-8">
      <style>
        .mystyle {
            font-size: 11pt; 
            font-family: Arial;
            border-collapse: collapse; 
            border: 1px solid silver;
        }
        .mystyle td, th {
            padding: 5px;
        }

        .mystyle tr:nth-child(even) {
            background: #E0E0E0;
        }

        .mystyle tr:hover {
            background: silver;
            cursor: pointer;
        }
        a {
            float:left;
            clear:left;
        }

        .mystyle table { width:100%; }
        .mystyle td, .mystyle th { width:3%; }
        .mystyle td + td, .mystyle th + th { width:5%; }
        .mystyle td + td + td, .mystyle th + th + th { width:5%; }     
        .mystyle td + td + td +td, .mystyle th + th + th +th { width:5%; }     
        .mystyle td + td + td +td+td, .mystyle th + th + th +th+th { width:5%; }     
        .mystyle td + td + td +td+td+td, .mystyle th + th + th +th+th+th { width:8%; }     
        .mystyle td + td + td +td+td+td+td, .mystyle th + th + th +th+th+th+th { width:11%; }     
        .mystyle td + td + td +td+td+td+td+td, .mystyle th + th + th +th+th+th+th+th { width:3%; }     

      </style>
      </head>

      <body>
      '''
    html_string_end = '''
      </body>
    </html>
    '''
    dir_list = os.listdir("CCP-perpetrtors")
    for file in dir_list:
        a = pd.read_excel(f"CCP-perpetrtors/{file}")
        pd.set_option('colheader_justify', 'center')  # FOR TABLE <th>



        # OUTPUT AN HTML FILE
        html = a.to_html(
                escape=False, render_links=True,
                na_rep='',
                classes='mystyle',index=False)
        file = file.replace("xlsx", "html")
        with open(f'output4_11/{file}', 'w') as f:
            f.write(html_string_start)
            f.write(html)
            f.write(html_string_end)

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
"""
https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.read_csv.html
https://stackoverflow.com/questions/52104682/rendering-a-pandas-dataframe-as-html-with-same-styling-as-jupyter-notebook

https://stackoverflow.com/questions/50807744/apply-css-class-to-pandas-dataframe-using-to-html
"""