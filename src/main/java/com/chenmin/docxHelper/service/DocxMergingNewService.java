package com.chenmin.docxHelper.service;

public interface DocxMergingNewService {
    void merge(String filename1, String filename2) throws Exception;

    void merge() throws Exception;
}
