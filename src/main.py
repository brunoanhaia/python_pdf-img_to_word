import os
from pdf2image import convert_from_path
from docx import Document
from docx.shared import Cm


def main():
    current_dir = os.getcwd()
    input_dir = os.path.join(current_dir, '..\\input')
    temp_dir = os.path.join(current_dir, '..\\temp')
    output_dir = os.path.join(current_dir, '..\\output')

    create_dirs(temp_dir, input_dir, output_dir)
    input_files = os.listdir(input_dir)

    for file in input_files:
        convert_file_to_pdf(file, temp_dir, input_dir, output_dir)

    create_docx(output_dir)


def create_docx(output_dir):
    document = Document()
    file_list = os.listdir(output_dir)
    for file in file_list:
        document.add_picture(os.path.join(output_dir, file), width=Cm(16), height=Cm(21))
        document.add_page_break()

    document.save('output.docx')


def create_dirs(temp_dir, input_dir, output_dir):
    if not os.path.isdir(temp_dir):
        os.mkdir(temp_dir)
    else:
        temp_dir_files = os.listdir(temp_dir)

        for file in temp_dir_files:
            try:
                os.remove(os.path.join(temp_dir, file))
            except OSError as e:
                print("Error: %s : %s" % (file, e.strerror))

    if not os.path.isdir(input_dir):
        os.mkdir(input_dir)

    if not os.path.isdir(output_dir):
        os.mkdir(output_dir)


def convert_file_to_pdf(input_file, temp_dir, input_dir, output_dir):
    input_file = os.path.join(input_dir, input_file)
    pil_images = convert_from_path(input_file, output_folder=temp_dir)
    base_filename = os.path.splitext(os.path.basename(input_file))[0]

    for page in pil_images:
        save_page(output_dir, base_filename, page)


def save_page(output_dir, base_filename, page):
    count = 0
    while os.path.isfile(os.path.join(output_dir, base_filename + '_' + str(
            count) + '.jpg')):
        count += 1

    img_filename = os.path.join(output_dir, base_filename + '_' +
                                str(count) + '.jpg')
    print(img_filename)
    page.save(img_filename, 'JPEG')


def get_folder_file_count(folder):
    dir_files = os.listdir(folder)
    return len(dir_files)


if __name__ == '__main__':
    main()
