class Node():
    def __init__(self, data = None):
        self.left = None
        self.right = None
        self.data = data

    def setLeft(self, node):
        self.left = node

    def setRight(self, node):
        self.right = node

    def getLeft(self):
        return self.left

    def getRight(self):
        return self.right

    def setData(self, data):
        self.data = data

    def getData(self):
        return self.data

def inorder(Tree):
    if Tree:
        inorder(Tree.getLeft())
        print(Tree.getData(), end = ' ')
        inorder(Tree.getRight())
    return

def preorder(Tree):
    if Tree:
        print(Tree.getData(), end = ' ')
        preorder(Tree.getLeft())
        preorder(Tree.getRight())
    return

def postorder(Tree):
    if Tree:
        postorder(Tree.getLeft())
        postorder(Tree.getRight())
        print(Tree.getData(), end = ' ')
    return

if __name__ == '__main__':
    root = Node(1)
    root.setLeft(Node(2))
    root.setRight(Node(3))
    root.left.setLeft(Node(4))

    while(True):
        choice = int(input("\n ------------------------------------------------ \nEnter your choice: \n 1.inorder() \n 2.preorder() \n 3.postorder \n 4.Exit \n ------------------------------------------------ \n"))
        if choice == 1:
            inorder(root)
        elif choice == 2:
            preorder(root)
        elif choice == 3:
            postorder(root)
        else:
            break