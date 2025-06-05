package com.yw.easy_excel_test.entity;

public class TransferYCheckResult {
    private int transferYSum;
    private int listSize;

    public TransferYCheckResult(int transferYSum, int listSize) {
        this.transferYSum = transferYSum;
        this.listSize = listSize;
    }

    public int getTransferYSum() {
        return transferYSum;
    }

    public int getListSize() {
        return listSize;
    }
}
