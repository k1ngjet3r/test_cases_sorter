flash_user = ['flash']
multi_user = ['multi', 'primary' 'secondary']
press_button = ['long press', 'short press', 'press "end" key']
user = ['guest', 'driver']
invalid = ['audiobook']

exceptions = [flash_user, multi_user, press_button, invalid]
expts = ['flash_user', 'multi_user', 'press_button', 'invalid']

exp_dict = {name: item for name, item in zip(expts, exceptions)}

for i in exp_dict:
    print(i)
