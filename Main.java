import java.util.*;

public class Main{
    public static void main(String[] args){
        int[] arr = {1, 2, 3, 4};
        int ans = subMax(arr);
        System.out.println(ans);
    }

    public static int reverse(int n){
        int rev
        while (n!=0){

        }
    }

    public static int subMin(int[] arr){
        int min = Integer.MAX_VALUE;
        int runningMin = 0;
        for (int i=0; i < arr.length; i++){
            runningMin = runningMin + arr[i];
            if (runningMin < min){
                min = runningMin;
            }
            if (runningMin > 0){
                runningMin = 0;
            }
        }
        return min;
    }

    public static int subMax(int[] arr){
        int max =Integer.MIN_VALUE, runningMax=0;

        for (int i=0; i < arr.length; i++){
            runningMax = runningMax+arr[i];

            if (runningMax > max){
                max = runningMax;
            }

            if (runningMax < 0){
                runningMax =0;
            }
        }
        return max;
    }

    public static void sortZeros(int[] arr){
        int start =0, end = arr.length-1;

        while (start < end){
            while(start < end && arr[start] ==0){
                start++;
            }

            while (start < end && arr[end] ==1){
                end--;
            }

            if (start < end){
                int temp = arr[start];
                arr[start] = arr[end];
                arr[end] = temp;
            }
        }
    }

    public static int[] common(int[] arr, int[] arr2){
        int unique=0;
        for (int i =0; i< arr.length;i++){
            for (int j=0;j<arr2.length;j++){
                if(arr[i] == arr2[j]){
                    arr[unique] = arr[i];
                    unique++;
                    break;
                }
            }
        }
        return Arrays.copyOf(arr,unique);
    }

    public static int[] mergeArrays(int[] arr, int[] arr2){
        int[] mixed = new int[arr.length+arr2.length];

        for (int i=0; i < arr.length; i++){
            mixed[i] = arr[i];
        }

        for (int i=0; i < arr2.length;i++){
            mixed[arr.length+i] = arr2[i];
        }

        return mixed;
    }
    public static int getAnyMax(int[] arr, int n){
        int[] ans = removDup(mergeSort(arr));
        if (n > ans.length){
            System.out.print("No max exists for the given index, hence returning ");
            return -1;
        }
        int max = ans[n-1];
        return max;
    }

    public static void secMax(int[] arr){
        int max = arr[0], sec_max = arr[1];

        for(int i=0; i < arr.length; i++){
            if(arr[i] > max){
                sec_max = max;
                max = arr[i];
            } else if(arr[i] > sec_max && arr[i] != max){
                sec_max = arr[i];
            }
        }
        if (sec_max != max){
            System.out.println(sec_max);
        } else{
            System.out.println("No second max exists");
        }
    }

    public static int[] mergeSort(int[] arr){
        if(arr.length <=1){
            return arr;
        }

        int mid = arr.length/2;
        int[] left = mergeSort(Arrays.copyOfRange(arr,0,mid));
        int[] right = mergeSort(Arrays.copyOfRange(arr,mid,arr.length));

        return mix(left,right);
    }

    public static int[] mix(int[] first, int[] sec){
        int[] mixed = new int[first.length + sec.length];
        int i=0, j=0,k=0;
        while (i < first.length && j< sec.length){
            mixed[k++] = first[i] < sec[j] ? first[i++] : sec[j++];
        }

        while (i < first.length){
            mixed[k++] = first[i++];
        }

        while(j < sec.length){
            mixed[k++] = sec[j++];
        }
        return mixed;
    }

    public static int[] removDup(int[] arr){
        int unique =0;
        for(int i=0; i< arr.length; i++){
            boolean isDup = false;
            for(int j=0; j<i;j++){
                if(arr[i]==arr[j]){
                    isDup = true;
                    break;
                }
            }
            if(!isDup){
                arr[unique++] = arr[i];
            }
        }
        return Arrays.copyOf(arr,unique);
    }

}