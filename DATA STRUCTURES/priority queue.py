class PriorityQueue():
    def __init__(self):
        self.queue = []

    def __str__(self):
        return ' '.join([str("priority: "+str(i[0])+ "  element: "+str(i[1])+"   ") for i in self.queue])

    def isEmpty(self):
        return len(self.queue) == 0

    def insert(self, data, priority):
        self.queue.append((priority, data))
        self.queue.sort(reverse=True)

 
    def delete(self):
        if(len(self.queue)<=0):
            print("queue underflow")
            return -1
        else:
            return self.queue.pop()
 

if __name__ == '__main__':
    n = int(input("enter number of elements in Queue: "))
    
    myQueue = PriorityQueue()
    for i in range(n):
        el = int(input("enter element in Queue: "))
        p = int(input("enter priority : "))
        myQueue.insert(el,p)
    print(myQueue)

    while(True):
        choice = int(input("\n ------------------------------------------------ \nEnter your choice: \n 1.print() \n 2.delete() \n 3.Exit \n ------------------------------------------------ \n"))
        if choice == 1:
            print(myQueue)
        elif choice == 2:
            myQueue.delete()
            print(myQueue)
        else:
            break