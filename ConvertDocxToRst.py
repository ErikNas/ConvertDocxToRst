import os
from PIL import Image  # Модуль Pillow для конвертирования картинок


current_dir = os.getcwd()
os.chdir(current_dir)
files = os.listdir(current_dir+'\\src')
docx_names = filter(lambda x: x.endswith('.docx'), files)

print(f'Start converting docx to rst from directory '+current_dir)

for docx_name in docx_names:

    # Конвертируем документы в RST
    # ToDo переписать вызов программы под библиотеку
    docx_name_without_ex = docx_name.split('.')[0]
    print(docx_name_without_ex)
    pandoc_command = ('pandoc --wrap=none --extract-media='+docx_name_without_ex
                      + '_image --resource-path='+docx_name_without_ex
                      + '_image/media -f docx src/'+docx_name_without_ex
                      + '.docx -t rst -o '+docx_name_without_ex+'.rst')

    os.system(pandoc_command)

    # Пробегаемся по картинками и при необходимости конвертим все в PNG
    need_chng_img_ex = False
    image_path = '\\'+docx_name_without_ex+'_image\\media\\'
    files_image = os.listdir(current_dir+image_path)
    image_names = filter(lambda x: x.endswith(('.wmf', '.gif')), files_image)
    for image_name in image_names:
        need_chng_img_ex = True
        image_name_without_ex = image_name.split('.')[0]
        image_file = Image.open(current_dir+image_path+image_name)
        image_file.save(current_dir+image_path+image_name_without_ex+'.png')
        os.remove(current_dir+image_path+image_name)

    # Если конвертировали картинки, то поправим ссылки на них в документе
    if need_chng_img_ex:
        with open(docx_name_without_ex+'.rst', 'r', encoding='utf-8') as file_rst:
            file_text = file_rst.read()

        file_text = file_text.replace(".wmf", ".png")
        file_text = file_text.replace(".gif", ".png")

        with open(docx_name_without_ex+'.rst', 'w', encoding='utf-8') as file_rst:
            file_rst.write(file_text)

print(f'Converting is over')
