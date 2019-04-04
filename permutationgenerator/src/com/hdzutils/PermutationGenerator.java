package com.hdzutils;

import java.util.ArrayList;

public class PermutationGenerator {
    //从generator中的字符选择出select个字符的排列
    public static ArrayList<String> permutation(String generator, int select) {
        ArrayList<String> result = new ArrayList<>();
        if (generator.length() < select) {
            return result;
        }

        if (select == 1) {
            for (int i=0; i < generator.length(); i++) {
                result.add(""+generator.charAt(i));
            }
            return result;
        }

        for (int i=0; i < generator.length(); i++) {
            StringBuffer sb = new StringBuffer(generator);
            char pickedChar = sb.charAt(i);
            sb.deleteCharAt(i);

            ArrayList<String> subEnums = permutation(sb.toString(), select-1);
            for (int j=0; j < subEnums.size(); j++) {
                result.add(pickedChar + subEnums.get(j));
            }
        }
        return result;
    }

    //生成generator中的字符的所有全排列
    public static ArrayList<String> fullPermutation(String generator) {
        ArrayList<String> result = new ArrayList<>();
        if (generator.length() == 0) {
            result.add(generator);
            return result;
        }

        for (int i=0; i < generator.length(); i++) {
            StringBuffer sb = new StringBuffer(generator);
            char pickedChar = sb.charAt(i);
            sb.deleteCharAt(i);

            ArrayList<String> subEnums = fullPermutation(sb.toString());
            for (int j=0; j < subEnums.size(); j++) {
                result.add(pickedChar + subEnums.get(j));
            }
        }
        return result;
    }

    public static void main(String[] args) {
	    // write your code here
        ArrayList<String> enums = fullPermutation("123456");
        for(int i = 0; i < enums.size(); i++) {
            System.out.println(enums.get(i));
        }

        ArrayList<String> enums1 = permutation("123456", 3);
        for(int i = 0; i < enums1.size(); i++) {
            System.out.println(enums1.get(i));
        }
    }
}
