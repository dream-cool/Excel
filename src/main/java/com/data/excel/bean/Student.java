package com.data.excel.bean;

/**
 * @author ：clt
 * @Date ：Created in 21:27 2019/12/31
 */
public class Student {
    private String stuNum;
    private String name;
    private String stuClass;
    private String IdCardInvalidTime;
    private String politicalOutlook;
    private String postalAddress;
    private String homeAddress;
    private String tel;
    private String email;

    public void setProperty(int i,String value){
        switch (i){
            case 1:
                this.setStuNum(value);
                break;
            case 2:
                this.setName(value);
                break;
            case 3:
                this.setStuClass(value);
                break;
            case 4:
                this.setIdCardInvalidTime(value);
                break;
            case 5:
                this.setPoliticalOutlook(value);
                break;
            case 6:
                this.setPostalAddress(value);
                break;
            case 7:
                this.setHomeAddress(value);
                break;
            case 8:
                this.setTel(value);
                break;
            case 9:
                this.setEmail(value);
                break;
            default:
                return;
        }
    }

    public String getStuNum() {
        return stuNum;
    }

    public void setStuNum(String stuNum) {
        this.stuNum = stuNum;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getStuClass() {
        return stuClass;
    }

    public void setStuClass(String stuClass) {
        this.stuClass = stuClass;
    }

    public String getIdCardInvalidTime() {
        return IdCardInvalidTime;
    }

    public void setIdCardInvalidTime(String idCardInvalidTime) {
        IdCardInvalidTime = idCardInvalidTime;
    }

    public String getPoliticalOutlook() {
        return politicalOutlook;
    }

    public void setPoliticalOutlook(String politicalOutlook) {
        this.politicalOutlook = politicalOutlook;
    }

    public String getPostalAddress() {
        return postalAddress;
    }

    public void setPostalAddress(String postalAddress) {
        this.postalAddress = postalAddress;
    }

    public String getHomeAddress() {
        return homeAddress;
    }

    public void setHomeAddress(String homeAddress) {
        this.homeAddress = homeAddress;
    }

    public String getTel() {
        return tel;
    }

    public void setTel(String tel) {
        this.tel = tel;
    }

    public String getEmail() {
        return email;
    }

    public void setEmail(String email) {
        this.email = email;
    }

    @Override
    public String toString() {
        return "Student{" +
                "stuNum='" + stuNum + '\'' +
                ", name='" + name + '\'' +
                ", stuClass='" + stuClass + '\'' +
                ", IdCardInvalidTime='" + IdCardInvalidTime + '\'' +
                ", politicalOutlook='" + politicalOutlook + '\'' +
                ", postalAddress='" + postalAddress + '\'' +
                ", homeAddress='" + homeAddress + '\'' +
                ", tel='" + tel + '\'' +
                ", email='" + email + '\'' +
                '}';
    }
}
