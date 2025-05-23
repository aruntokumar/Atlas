package Searching;

public class Binary {
    public static void main(String[] args) {
        // int[] arr = {122,113,105,97,88,79,52,44,36,28,19};
//        int[] arr = {19,28,36,44,52,79,88,97,105,113,122};
        char[] arr = {'c', 'f','j'};
        char target = 'c';
//        System.out.println(binary(arr,target));
        System.out.println(nextGreatestLetter(arr,target));
    }

    public static int binary(int[] arr, int target) {
        int start = 0;
        int end = arr.length-1;

        boolean isAsc = arr[start] < arr[end];

        while (start <= end) {
            int mid = start + (end-start) /2;

            if (target == arr[mid]) {
                return mid;

            } else if (isAsc) {
                if (target > arr[mid]) {
                    start = mid+1;
                } else if (target < arr[mid]) {
                    end = mid-1;
                }
            } else {
                if (target < arr[mid]) {
                    start = mid+1;
                } else if (target > arr[mid]) {
                    end = mid-1;
                }
            }
        }
        return -1;
    }

    static char nextGreatestLetter(char[] letters, char target) {
        int start = 0;
        int end = letters.length - 1;

        while (start <= end) {
            int mid = start + (end - start) /2;
            if (target < letters[mid]) {
                end = mid-1;
            } else {
                start = mid+1;
            }
        }

        return letters[start % letters.length];
    }
}
