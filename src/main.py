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

    temp_files = []

    for file in input_files:
        temp_files += pdf_to_img(file, temp_dir, input_dir, output_dir)

    create_img_docx(temp_files, output_dir)


def create_img_docx(file_list, output_dir):
    document = Document()

    for file in file_list:
        document.add_picture(file, width=Cm(15))
        document.add_page_break()

    for section in document.sections:
        section.top_margin = Cm(3)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(3)
        section.right_margin = Cm(2)

    document.save(output_dir + '\\output.docx')


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


def pdf_to_img(input_file, temp_dir, input_dir, output_dir):
    input_file = os.path.join(input_dir, input_file)
    images_path = convert_from_path(input_file, output_folder=temp_dir, fmt='png', paths_only=True)

    return images_path


# TODO: Create function to avoid output file replacement
def check_existing_files(input_dir, base_filename):
    count = 0
    while os.path.isfile(os.path.join(input_dir, base_filename + '_' + str(
            count) + '.jpg')):
        count += 1

    img_filename = os.path.join(input_dir, base_filename + '_' +
                                str(count) + '.jpg')
    print(img_filename)


if __name__ == '__main__':
    main()
