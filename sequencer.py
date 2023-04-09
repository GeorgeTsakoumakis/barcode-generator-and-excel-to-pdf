import single

def sequencer():
    start = int(input('Enter starting number: \n> '))
    end = int(input('Enter ending number: \n> '))
    filename, copy, cell = single.main()
    for i in range(start, end+1):
        data = 'ATL-00'+str(i)
        single.barcode_generator(data)
        single.pic_in_excel(data, filename, copy, cell)
        single.xl_to_pdf(data, copy)

if __name__ == '__main__':
    sequencer()
