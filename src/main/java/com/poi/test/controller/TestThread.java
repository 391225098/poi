package com.poi.test.controller;

import java.util.concurrent.Executor;
import java.util.concurrent.Executors;

/**
 * Created with IntelliJ IDEA.
 * User: wb252654
 * Date: 17-9-13
 * Time: 下午4:38
 * To change this template use File | Settings | File Templates.
 */
public class TestThread {

    private static Executor executor = Executors.newCachedThreadPool();
    private static Executor executor1 = Executors.newFixedThreadPool(10);
    private static Executor executor2 = Executors.newSingleThreadExecutor();
    private static Executor executor3 = Executors.newScheduledThreadPool(10);

    public static void main(String[] args) {
        System.out.println(System.nanoTime());
        for (int i = 0; i < 200; i++) {
//            executor.execute(new Task());
//            executor1.execute(new Task());
//            executor2.execute(new Task());
            executor3.execute(new Task());

        }
        System.out.println(System.nanoTime());
    }

    static  class Task implements Runnable {

        @Override
        public void run() {
            System.out.println(Thread.currentThread().getName());
//            throw new RuntimeException("lihaile");
        }
    }
}
