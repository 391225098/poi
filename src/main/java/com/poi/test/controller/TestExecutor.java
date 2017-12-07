package com.poi.test.controller;

import java.util.concurrent.*;

/**
 * Created with IntelliJ IDEA.
 * User: wb252654
 * Date: 17-9-14
 * Time: 上午9:36
 * To change this template use File | Settings | File Templates.
 */
public class TestExecutor {

    private static ExecutorService e = Executors.newFixedThreadPool(10);

    public static void main(String[] args) {
        Task task = new Task();
        Future<Boolean> future = e.submit(task);
        System.out.println("do something");
        System.out.println(System.nanoTime());
        try {
            Boolean result = future.get();
            System.out.println(result);
            System.out.println(System.nanoTime());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    static class Task implements Callable<Boolean> {

        @Override
        public Boolean call() throws Exception {
            boolean result = true;
            try {
                TimeUnit.SECONDS.sleep(2);
                throw new RuntimeException("123");
            } catch (Exception e) {
                e.printStackTrace();
                result = false;
            }
            return result;
        }
    }
}
