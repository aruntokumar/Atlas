package Sorting;

public class Selection {
    public static void main(String[] args) {
        int[] arr = {-4,-6,-7,4,12,23,56,67,89,99}; //9
        int target = 23;
        System.out.println(binary(arr,target));
    }

    public static int binary(int[] arr, int target) {
        int start = 0;
        int end = arr.length-1;

        while (start <= end) {
            int mid = start + (end-start)/2; //4
            if (target > arr[mid]) {
                start = mid+1;
            } else if (target < arr[mid]) {
                end = mid-1;
            } else {
                return mid;
            }
        }
        return -1;
    }
}