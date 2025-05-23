package LinkedList;

public class LinkedList {
    Node head;
    Node tail;
    int size;

    public class Node {
        Object val;
        Node next;

        public Node(Object val){
            this.val = val;
        }
    }

    public LinkedList(){
        this.size = 0;
    }

    public void add(Object val) {
        Node node = new Node(val);
        node.next = head;
        head = node;

        if (tail == null){
            tail = head;
        }
        size++;
    }

    public void append(Object val){
        Node node = new Node(val);
        if (tail == null){
            add(val);
            return;
        }
        tail.next = node;
        tail = node;
        size++;
    }

    public void display() {
        Node temp = head;
        System.out.print("LinkedList: ");
        while (temp != null){
            System.out.print(temp.val + " -> ");
            temp = temp.next;
        }
        System.out.println("End");
    }

    public void size(){
        System.out.println("size: " + size);
    }

    public static void main(String[] args){
        LinkedList list = new LinkedList();
        list.append(3);
        list.append(5);
        list.append(7);
        list.append(9);
        list.display();
        list.size();
    }
}
