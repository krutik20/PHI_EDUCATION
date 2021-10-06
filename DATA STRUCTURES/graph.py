
class AdjacencyList():
    def __init__(self):
        self.List = {}

    def addEdge(self, fromVertex, toVertex):
        if fromVertex in self.List.keys():
            self.List[fromVertex].append(toVertex)
        else:
            self.List[fromVertex] = [toVertex]

    def printList(self):
        for i  in self.List:
            print(i,'->',' -> '.join([str(j) for j in self.List[i]]))

if __name__ == '__main__':
    al = AdjacencyList()

    e = int(input("Enter number of edges in the graph"))

    for i in range(e):
        from1=int(input("Enter From which vertex : "))
        to=int(input("To which vertex : "))
        al.addEdge(from1, to)

    al.printList()
