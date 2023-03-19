package com.chenmin.docxHelper.service;

import java.io.FileNotFoundException;
import java.util.List;

public interface DocxMergingNewService {
    void merge(String filename1, String filename2) throws Exception;

    void merge() throws Exception;
}
