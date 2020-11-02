def matcher(keywords, sentance):
    for key in keywords:
        num_slices = int(len(sentance)) + 1 - int(len(key))
        for i in range(num_slices):
            if sentance[i: i + len(key)] == key:
                return True
    return False


def matcher_2(keywords, sentance):
    word_list = sentance.split()
    for key in keywords:
        if key in word_list:
            return True
    return False


sen = 'hello, my name is Pair'
keywords = ['I', 'air']

print(matcher(keywords, sen))

print(matcher_2(keywords, sen))
