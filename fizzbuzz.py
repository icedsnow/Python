count = 1

while count < 100:
    if count % 3 == 0 and count % 5 == 0:
        print('FizzBuzz')
        count += 1
    elif count % 3 == 0 and not count % 5 == 5 and not count == 3:
            print('Fizz')
            count += 1
    elif count % 5 == 0 and not count % 3 == 0 and not count == 5:
        print('Buzz')
        count += 1
    else:
        print(count)
        count +=1
        
    
