package com.poi.test.controller;

import java.util.HashMap;
import java.util.UUID;

/**
 * Created with IntelliJ IDEA.
 * User: wb252654
 * Date: 17-9-14
 * Time: 下午4:23
 * To change this template use File | Settings | File Templates.
 */
public class TestMap {
    public static void main(String[] args) {
        final HashMap<String, String> map = new HashMap<String, String>(2);
        for (int i = 0; i < 10000; i++) {
            new Thread(new Runnable() {
                @Override
                public void run() {
                    map.put(UUID.randomUUID().toString(), "");
                }
            }).start();
        }
    }
}
