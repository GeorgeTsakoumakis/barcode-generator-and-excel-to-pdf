import sequencer
import single
from time import sleep

choice = input('Would you like to create:\n1) multiple pdfs or \n2) a single pdf?\nAnswer with 1/2: ')
if choice == '1':
    sequencer.sequencer()
elif choice == '2':
    single.run()
else:
    print('Invalid input. Please try again.')
    sleep(2)