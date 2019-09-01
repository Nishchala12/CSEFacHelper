package application;

public class FireData
{
	public String att, perc, sub, subs;
	
	public FireData() {}
	
    public String getAtt() { return att; }

    public String getPerc() { return perc; }
    
    public String getSub() { return sub; }
    
    public String getSubs() { return subs; }
    
    public void setAtt(String name) {
        this.att = name;
    }
    
    public void setPerc(String name) {
        this.perc = name;
    }
    
    public void setSub(String name) {
        this.sub = name;
    }
    
    public void setSubs(String name) {
        this.subs = name;
    }
}