def matcher(keywords, sentance):
    for key in keywords:
        num_slices = int(len(sentance)) + 1 - int(len(key))
        for i in range(num_slices):
            if sentance[i: i + len(key)] == key:
                return True
    return False


sen = 'hello, my name is jeter'
keywords = ['I', 'you']

print(matcher(keywords, sen))