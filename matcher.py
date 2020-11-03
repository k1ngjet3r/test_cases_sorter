import re


def matcher(keywords, sentance):
    for key in keywords:
        num_slices = int(len(sentance)) + 1 - int(len(key))
        for i in range(num_slices):
            if sentance[i: i + len(key)] == key:
                return True
    return False


def matcher_split(keywords, sentance):
    clean_sentance = re.sub(r'[^\w]', ' ', sentance.lower())
    word_list = clean_sentance.split()
    for key in keywords:
        if key in word_list:
            return True
    return False


sen = '1.hello! My name is Pair.'
keywords = ['my', 'air']

print(matcher_split(keywords, sen))
