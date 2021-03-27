def txt_2_list(filename):
    txt_file = open(filename, 'r')
    return filename, txt_file.split('\n')

