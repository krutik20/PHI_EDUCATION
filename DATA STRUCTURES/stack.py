class Stack():
    def __init__(self, stacksize = 10):
        self.stack = []
        self.stacksize = stacksize

    def __str__(self):
        return ' '.join([str(i) for i in self.stack])

    def push(self, data):
        if len(self.stack) >= self.stacksize:
            print('Stack Overflow')
        else:
            self.stack.append(data)

    def pop(self):
        if len(self.stack) <= 0:
            print("stack underflow")
            return -1
        else:
            return self.stack.pop()

    def peek(self):
        if len(self.stack) <= 0:
            return -1
        else:
            return self.stack[len(self.stack) - 1]

    def isEmpty(self):
        return self.stack == []

    def size(self):
        return len(self.stack)

if __name__ == '__main__':
    n = int(input("enter number of elements in stack: "))
    
    myStack = Stack(n)
    for i in range(n):
        el = int(input("enter element in stack: "))
        myStack.push(el)
    print(myStack)
    while(True):
        choice = int(input("Enter your choice: 1.push() \n 2.pop() \n 3.peek() \n 4.print() \n 5.exit"))
        if choice == 1:
            el = int(input("enter element in stack: "))
            myStack.push(el)
        elif choice == 2:
            myStack.pop()
        elif choice == 3:
            print(myStack.peek())
        elif choice == 4:
            print(myStack)
        else:
            break