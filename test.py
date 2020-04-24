import zipfile
import os
from shutil import copyfile


orig_dir_path = 'C:/temp/docx/org/' #Оригиналы
extracted_dir_path = 'C:/temp/docx/extracted/' #распакованные
ready_for_copy_dir_path = 'C:/temp/docx/ready_for_copy/'
final_dir_path = 'C:/temp/docx/final/'

single_example = 'C:/temp/docx/example/doc1.docx'
doc_sufix = '/word/document.xml'


name_without_extension_list = []

#Распаковываем оригинальные файлы
original_files = os.listdir(orig_dir_path)
for f in original_files: #Распаковываем    
    splitted_file_name = os.path.splitext(str(f)) #Убираем extension    
    name_without_extension = splitted_file_name[0] 
    name_without_extension_list.append(name_without_extension)
    
    with zipfile.ZipFile(orig_dir_path+f,"r") as zip_ref:
        zip_ref.extractall(extracted_dir_path+name_without_extension)
    print(name_without_extension + ' extracted')

#Распаковываем единственный образец для каждой папки из образца
with zipfile.ZipFile(single_example,"r") as zip_ref:
    for f in name_without_extension_list:       
        zip_ref.extractall(ready_for_copy_dir_path+f)
print('Образцы распакованы')
    
#Копируем файлы из исходных данных в каталог
for f in name_without_extension_list:
    src = extracted_dir_path + f + doc_sufix
    dst = ready_for_copy_dir_path + f + doc_sufix
    copyfile(src, dst)
    print(f + " Успешно скопирован")
        
#пакуем в zip
for f in name_without_extension_list:
    os.chdir(ready_for_copy_dir_path + f)    
    z = zipfile.ZipFile(final_dir_path + f + '.docx', 'w')        # Создание нового архива
    for root, dirs, files in os.walk('.'): # Список всех файлов и папок в директории (текущей)
        for file in files:
            z.write(os.path.join(root,file))         # Создание относительных путей и запись файлов в архив
    z.close()
    print("Файл успешно преобразован в docx")
