package Searching;

public class Linear {
    public static void main(String[] args) {
        int[] arr = {23,13,455,32,39};
        int target = 23;
        System.out.println(linear(arr,target));
    }

    static boolean linear(int[] arr, int target) {
        for (int i = 0; i< arr.length-1;i++) {
            if (arr[i] == target) {
                return true;
            }
        }
        return false;
    }

    public static void selection(int[] arr){
        int n = arr.length;
        for (int i=0; i < n-1; i++){
            int min = i;
            for (int j = i+1; j <n;j++){
                if (arr[j] > arr[min]) min = j;
            }
            int temp = arr[min];
            arr[min] = arr[i];
            arr[i] = temp;
        }
    }

    public static void insertion(int[] arr){
        for (int i=0; i<arr.length-1;i++){
            for (int j = i+1; j>0;j--){
                if (arr[j]<arr[j-1]){
                    int temp = arr[j];
                    arr[j] = arr[j-1];
                    arr[j-1] = temp;
                }else break;
            }
        }
    }
}
