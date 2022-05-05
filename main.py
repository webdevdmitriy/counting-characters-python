import docx, os, pypandoc, re



# Получаем все файлы ворд из папки docx. Преобразуем в txt и кладем в папку txt
#for filename in os.listdir("docx"):
#    pypandoc.convert_file(f"docx/{filename}", 'plain', outputfile=f"txt/{filename.replace('.docx', '')}.txt")



# Подсчитываем кол-во символов в файлах
file = open("result.txt", "w")
for filename in os.listdir("txt"):
   with open(os.path.join("txt", filename), 'r', encoding='utf-8') as f:
       data = f.read()
       data_format =  re.sub(r"\s\s+|\n|-", "", data)
       print(data_format)
       print(filename)
       result = f"{filename} {len(data_format)}\n"
       file.write(result)

