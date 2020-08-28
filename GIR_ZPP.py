#!/usr/bin/env python
# coding: utf-8

# In[75]:


#!/usr/bin/env python
# coding: utf-8
def count_sum_of_probs_gir_zpp():
    """
        1) Данная программка принимает на вход все множество эксель-файлов,
        парсит из них сведения, формирует сводную таблицу, в которой название столбца
        соответствует названию файла, из которого эти сведения получены. При этом,
        распарсиваются только вторые столбцы таблицы ("количество проб"). Формируется датафрейм
        где последний столбец - сумма по строкам.
        Далее полученный датафрейм сохранятеся в новый файл, именуемый 'свод_по_кол-ву_проб.xlsx'
        и сохраняется в текущей директории.

        2) Анализ за неделю - в стадии разработки...
        Также все множество эксель-файлов...
        Количество проб, отобранных за неделю, из них - не соответствуют, из несоответсвующих
        пищевки - столько-то, непищевки - столько-то...
    """

    import pandas as pd
    import os

    #######

    # ФУНКЦИИ!!!

    ###
    def check_wd_files_extension(list_of_files):
        """
        принимает на вход список файлов директории
        данная функция возвращает True в том случае, если
        все файлы директории имеют расширение .xls или .xlsx
        """
        flag = True
        for file in list_of_files:
            if (file.endswith('.xls') or
            file.endswith('.xlsx')):
                continue
            else:
                flag = False
                print("Обнаружен файл с недопустимым расширением!")
        return flag

    ###

    def check_first_sheet_in_files(list_of_files):
        """
        парсер получает на вход список файлов
        на извлечение данных по количеству отобранных проб
        проверяет названия первых листов.
        возвращает True если названия всех первых листов одинаковых
        """
        list_name_flag = True
        counter = 0
        prev_list_name = ''
        for file in list_of_files:
            data = pd.ExcelFile(file)
            my_sheet_names = data.sheet_names

            if prev_list_name != my_sheet_names[0] and counter > 0:
                list_name_flag = False
                print("Названия листов различаются в файлах!!!")
                print("Название первого листа в книге: ", prev_list_name)
                print("Название первого листа в книге: ", my_sheet_names[0])
                print("Название файла: ", file)
            prev_list_name = my_sheet_names[0]
            counter += 1
        if list_name_flag == False:
            print("проверьте указанные файлы!!!")
        return list_name_flag  

    ###

    def check_template(list_of_files):
        """
        парсер получает на вход список файлов
        на извлечение данных по количеству отобранных проб
        возвращает True если размерность ДФ (N строк; N столбцов) всех шаблонов равна
        """
        set_raws = set()
        set_col = set()
        file_counter = 0
        for file in list_of_files:
            file_counter += 1
            data = pd.ExcelFile(file)
            my_sheet_names = data.sheet_names
            current_df = data.parse(my_sheet_names[0])
            set_raws.add(current_df.shape[0])
            set_col.add(current_df.shape[1])
            #print(f'в {file_counter}-ой таблице, расположенной в {file},\
     #количество строк составило {current_df.shape[0]}, \
     #а столбцов - {current_df.shape[1]}')
        if len(set_raws)+len(set_col) == 2 and len(set_raws) == 1:
            print(f'шаблоны таблицы одинаковые как по количеству строк {set_raws.pop()}, так и - столбцов {set_col.pop()}')
            return True
        if len(set_raws) != 1:
            print(f'количество строк различается в шаблонах {set_raws}')
        if len(set_col) != 1:
            print(f'количество столбцов различается в шаблонах {set_col}')
        return False


    ###

    def parser_df(list_of_files):
        """
        парсер получает на вход список файлов
        на извлечение данных по количеству отобранных проб
        склеивает все столбцы в один фрейм данным и возвращает его
        предварительно проверив на размерность ДФ, названия его столбцов и строк
        и т.д.
        """
        data = pd.ExcelFile(list_of_files[0])
        my_sheet_names = data.sheet_names
        df = data.parse(my_sheet_names[0])
        df_summary = df.iloc[14:36,:2]
        file_counter = 0

        for file in list_of_files[1:]:
            file_counter += 1
            data = pd.ExcelFile(file)
            my_sheet_names = data.sheet_names
            current_df = data.parse(my_sheet_names[0])
            current_column = current_df.iloc[14:36,1:2]
            current_column[file] = current_df.iloc[14:36,1:2]
            df_summary = df_summary.join(current_column.iloc[:,-1], how='right')
            #print(current_df.iloc[14:36,:2])
            #pd.merge(df_summary,
                     #current_df.iloc[14:36,:2],
                     #how='inner',
                     #on='Unnamed: 0')
        # заменить все пропущенные значения на 0
        df_summary = df_summary.fillna(0)
        df_summary['СУММА'] = df_summary.sum(axis=1)
        # вернуть суммарный ДФ
        return df_summary     


    ######

    # этап предподготовки директории
    # количество файлов в ней
    curr_wd = os.getcwd()
    list_files = os.listdir()
    count_files_wd = len(list_files)
    print(f"Количество файлов в рабочей директории {curr_wd} составляет {count_files_wd} шт.")
    print()
    # проверка типа файлов
    all_files_excel = check_wd_files_extension(list_files)
    print('Отчет о проверке расширений файлов в директории: ', all_files_excel)
    print()

    # этап распарсивания файлов
    #проверка названия первого листа в файлах
    check_first_sheet_in_files(list_files)
    #проверка размерности
    template_ok = check_template(list_files)
    #непосредственно само распарсивание
    if template_ok:
        df = parser_df(list_files)
    else:
        print("Уппсссс, по всей видимости, что-то пошло не так...")
        input()
    print("Размерность полученного ДФ составила (строк, столбцов): ", df.shape)
    print()
    # Выкатываем объединенную табличку!
    df.to_excel('ГИР_ЗПП_табличка_с_суммой_по_всем_файлам.xlsx', startrow=0, index=False)
    input()

count_sum_of_probs_gir_zpp()

