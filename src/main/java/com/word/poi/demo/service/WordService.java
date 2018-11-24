package com.word.poi.demo.service;

import java.io.InputStream;
import java.io.OutputStream;

public interface WordService {
    void exportWord(InputStream in, OutputStream out, String city);
}
