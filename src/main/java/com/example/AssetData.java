package com.example;

public class AssetData {
	private Object assetNumber;
	private Object assetName;
	private Object purchaseDate;
	private Object purchasePrice;
	private Object AssetType;
	private Object bookRate;
	private Object closingBookRate;
	private Object bookAccumulatedDepreciation;
	public Object getBookAccumulatedDepreciation() {
		return bookAccumulatedDepreciation;
	}
	public void setBookAccumulatedDepreciation(Object bookAccumulatedDepreciation) {
		this.bookAccumulatedDepreciation = bookAccumulatedDepreciation;
	}
	public Object getClosingBookRate() {
		return closingBookRate;
	}
	public void setClosingBookRate(Object closingBookRate) {
		this.closingBookRate = closingBookRate;
	}
	public Object getBookRate() {
		return bookRate;
	}
	public void setBookRate(Object bookRate) {
		this.bookRate = bookRate;
	}
	public Object getAssetNumber() {
		return assetNumber;
	}
	public void setAssetNumber(Object assetNumber) {
		this.assetNumber = assetNumber;
	}
	public Object getAssetName() {
		return assetName;
	}
	public void setAssetName(Object assetName) {
		this.assetName = assetName;
	}
	public Object getPurchaseDate() {
		return purchaseDate;
	}
	public void setPurchaseDate(Object purchaseDate) {
		this.purchaseDate = purchaseDate;
	}
	public Object getPurchasePrice() {
		return purchasePrice;
	}
	public void setPurchasePrice(Object purchasePrice) {
		this.purchasePrice = purchasePrice;
	}
	public Object getAssetType() {
		return AssetType;
	}
	public void setAssetType(Object assetType) {
		AssetType = assetType;
	}
	
}
