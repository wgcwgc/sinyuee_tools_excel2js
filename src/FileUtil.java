/*
 * @(#)FileUtil.class 0.100 2009-2-20
 *
 * Copyright 2007 Smarch.
 * All rights reserved.
 */


import java.io.ByteArrayOutputStream;
import java.io.DataInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;

/**
 * @author runasea
 * @version 0.1.0.1, 2009-2-20
 * @since SMARCH0.1
 */
public class FileUtil {

  /**
   * 读取某个文件。
   * @param f 文件
   * @return 内容
   */
  public static byte[] readFile(final File f) {
    byte[] data = new byte[(int) f.length()];
    DataInputStream dis;
    try {
      dis = new DataInputStream(
          new FileInputStream(f));
      dis.readFully(data);
      dis.close();
      return data;
    } catch (IOException e) {
      e.printStackTrace();
    }
    return null;
  }

  /**
   * 从流中读取文件。
   * @param is 文件流
   * @return 内容
   */
  public static byte[] readFile(final InputStream is) {
    ByteArrayOutputStream baos = new ByteArrayOutputStream();
    DataInputStream dis = new DataInputStream(is);
    try {
      while (dis.available() > 0) {
        byte[] data = new byte[dis.available()];
        dis.readFully(data);
        baos.write(data);
      }
    } catch (IOException e) {
      e.printStackTrace();
    }
    return baos.toByteArray();
  }

  /**
   * 输出某个文件。
   * @param f 文件
   * @param data 内容
   */
  public static void writeFile(final File f, final byte[] data) {
    if (f.exists()) f.delete();
    FileOutputStream fos;
    try {
      fos = new FileOutputStream(f);
      fos.write(data);
      fos.close();
    } catch (IOException e) {
      e.printStackTrace();
    }
  }

  /**
   * 读取某个文本文件。
   * @param f 文件
   * @return 文本内容
   */
  public static String readFileText(final File f) {
    try {
      return new String(readFile(f), "utf-8");
    } catch (UnsupportedEncodingException e) {
      e.printStackTrace();
    }
    return null;
  }

  /**
   * 从流中读取某个文本文件。
   * @param is 文件流
   * @return 文本内容
   */
  public static String readFileText(final InputStream is) {
    return new String(readFile(is));
  }

  /**
   * 输出某个文本文件。
   * @param f 文件
   * @param s 文本
   */
  public static void writeFileText(final File f, final String s) {
    writeFile(f, s.getBytes());
  }

  /**
   * 拷贝文件。
   * @param src 源
   * @param dest 目标
   */
  public static void copyFile(File src, File dest) {
    writeFile(dest, readFile(src));
  }

  /**
   * 读取带有注释的文件。
   * @param f 文件
   * @return 去掉注释部分的文本
   */
  public static String readCommitFile(final File f) {
    String s = new String(readFile(f));
    String tmp = "";
    boolean commit = false;
    for (int i = 0; i < s.length(); i++) {
      char ch = s.charAt(i);
      if (ch == '#') {
        commit = true;
      }
      if (commit) {
        if (ch == '\r' || ch == '\n')
          commit = false;
      } else {
        tmp += ch;
      }
    }
    return tmp;
  }

  public static void main(final String[] cmd) {
    System.out.println(readCommitFile(new File("doc/api_trans.txt")));
  }

}
