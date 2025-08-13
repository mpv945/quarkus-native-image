package com.ewancle.model;

public class UploadResponse {
    public String originalName;
    public String storedName;

    public long size;

    public UploadResponse() {}
    public UploadResponse(String originalName, String storedName, long size) {
        this.originalName = originalName;
        this.storedName = storedName;
        this.size = size;
    }
}
