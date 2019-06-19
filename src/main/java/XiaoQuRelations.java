import java.io.Serializable;

public class XiaoQuRelations implements Serializable {

    // excel表中的字段
    private int id;
    private String city;
    private String district;
    private String block;
    private String type;
    private String company;
    private String jingdui;
    private String note;
    private String jingdui_type;
    private String jingdui_district;

    public int getId() {
        return id;
    }

    public String getCity() {
        return city;
    }

    public String getDistrict() {
        return district;
    }

    public String getBlock() {
        return block;
    }

    public String getType() {
        return type;
    }

    public String getCompany() {
        return company;
    }

    public String getJingdui() {
        return jingdui;
    }

    public String getNote() {
        return note;
    }

    public String getJingdui_type() {
        return jingdui_type;
    }

    public String getJingdui_district() {
        return jingdui_district;
    }

    public void setId(int id) {
        this.id = id;
    }

    public void setCity(String city) {
        this.city = city;
    }

    public void setDistrict(String district) {
        this.district = district;
    }

    public void setBlock(String block) {
        this.block = block;
    }

    public void setType(String type) {
        this.type = type;
    }

    public void setCompany(String company) {
        this.company = company;
    }

    public void setJingdui(String jingdui) {
        this.jingdui = jingdui;
    }

    public void setNote(String note) {
        this.note = note;
    }

    public void setJingdui_type(String jingdui_type) {
        this.jingdui_type = jingdui_type;
    }

    public void setJingdui_district(String jingdui_district) {
        this.jingdui_district = jingdui_district;
    }

    public void show(){
        System.out.println("id:"+id+"\n"+
                "city:"+city+"\n"+
                "district:"+district+"\n"+
                "block:"+block+"\n"+
                "type:"+type+"\n"+
                "company:"+company+"\n"+
                "jingdui:"+jingdui+"\n"+
                "note:"+note+"\n"+
                "jingdui_type:"+jingdui_type+"\n"+
                "jingdui_district:"+jingdui_district);
    }
}
