class Queue():
    def __init__(self, limit = 10):
        self.queue = []
        self.front = None
        self.rear = None
        self.limit = limit
        self.size = 0

    def __str__(self):
        return ' '.join([str(i) for i in self.queue])

    def isEmpty(self):
        return self.size <= 0

    def enqueue(self, data):
        if self.size >= self.limit:
            print("queue overflow")
            return -1    
        else:
            self.queue.append(data)

        if self.front is None:
            self.front = self.rear = 0
        else:
            self.rear = self.size

        self.size += 1

    def dequeue(self):
        if self.isEmpty():
            print("queue underflow")
            return -1          
        else:
            self.queue.pop(0)
            self.size -= 1
            if self.size == 0:
                self.front = self.rear = 0
            else:
                self.rear = self.size - 1

    def getSize(self):
        return self.size

    def peek(self):
        if len(self.queue) <= 0:
            print("queue is empty")
            return -1
        else:
            return self.queue[self.front]

if __name__ == '__main__':
    n = int(input("enter number of elements in Queue: "))
    
    myq = Queue(n)
    for i in range(n):
        el = int(input("enter element in Queue: "))
        myq.enqueue(el)
    print(myq)

    while(True):
        choice = int(input("\n ------------------------------------------------ \nEnter your choice: \n 1.enqueue() \n 2.dequeue() \n 3.peek() \n 4.print() \n 5.size \n 6.Exit \n ------------------------------------------------ \n"))
        if choice == 1:
            el = int(input("enter element in Queue: "))
            myq.enqueue(el)
        elif choice == 2:
            myq.dequeue()
        elif choice == 3:
            print(myq.peek())
        elif choice == 4:
            print(myq)
        elif choice == 5:
            print('Queue Size:',myq.getSize())
        else:
            break