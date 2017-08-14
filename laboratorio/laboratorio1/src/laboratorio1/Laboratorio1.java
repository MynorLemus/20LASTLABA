/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package laboratorio1;
import java.util.Scanner;
public class Laboratorio1 {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
      int a;
      System.out.println("ingresar primer numero");
      Scanner g = new Scanner (System.in);
      a = g.nextInt();
      
      if (a%2==0){
        System.out.println(a +" es par");
      }
      
      else{
      System.out.println(a +" es impar");
      }
    }
    
}
