import java.util.Arrays;
import  java.util.Scanner;

public class Array2d{
    public static void main(String[] args){
        int[][] arr2 = new int[3][3];
        System.out.println("Enter array numbers: ");
        inp_arr(arr2);
        for (int i = 0; i < arr2.length;i++) {
            System.out.println(Arrays.toString(arr2[i]));
        }
        // we can enhance the above loop like below

//        for (int[] a: arr2) {
//            System.out.println(Arrays.toString(a));
//        }
    }

    // 2D array

    static void inp_arr(int[][] arr) {
        Scanner inp = new Scanner(System.in);
        for (int i=0; i < arr.length; i++) {
            for (int j=0; j < arr[i].length;j++) {
                arr[i][j] = inp.nextInt();
            }
        }
    }

}