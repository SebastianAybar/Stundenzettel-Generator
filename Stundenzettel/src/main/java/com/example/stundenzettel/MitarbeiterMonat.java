package com.example.stundenzettel;

import org.apache.poi.ss.usermodel.Cell;

public class MitarbeiterMonat {

    private String svBrutto;
    private String abrechnungsmonat;
    private String mitarbeiternummer;
    private String nachnameVorname;
    /*private String nr;
    private String berater;
    private String mandant;
    private String pgr;
    private String stkl;
    private String antyp;
    private String bschluesselKv;
    private String bschluesselRv;
    private String bschluesselAv;
    private String bschluesselPv;
    private String nameBetrieb;
    private String strassePostfach;
    private String plz;
    private String ort;
    private String svTage;
    private String natKennzeichen;
    private String midijobregelung;*/

    public String getAbrechnungsmonat() {
        return abrechnungsmonat;
    }
    public void setAbrechnungsmonat(String abrechnungsmonat) {
        this.abrechnungsmonat = abrechnungsmonat;
    }
    public String getMitarbeiternummer() {
        return mitarbeiternummer;
    }
    public void setMitarbeiternummer(String mitarbeiternummer) {
        this.mitarbeiternummer = mitarbeiternummer;
    }
    public String getNachnameVorname() {
        return nachnameVorname;
    }
    public void setNachnameVorname(String nachnameVorname) {
        this.nachnameVorname = nachnameVorname;
    }
    public String getSvBrutto() {
        return svBrutto;
    }
    public void setSvBrutto(String svBrutto) {
        this.svBrutto = svBrutto;
    }

    /*public String getNr() {
        return nr;
    }

    public void setNr(String nr) {
        this.nr = nr;
    }

    public String getBerater() {
        return berater;
    }

    public void setBerater(String berater) {
        this.berater = berater;
    }

    public String getMandant() {
        return mandant;
    }

    public void setMandant(String mandant) {
        this.mandant = mandant;
    }

    public String getPgr() {
        return pgr;
    }

    public void setPgr(String pgr) {
        this.pgr = pgr;
    }

    public String getStkl() {
        return stkl;
    }

    public void setStkl(String stkl) {
        this.stkl = stkl;
    }

    public String getAntyp() {
        return antyp;
    }

    public void setAntyp(String antyp) {
        this.antyp = antyp;
    }

    public String getBschluesselKv() {
        return bschluesselKv;
    }

    public void setBschluesselKv(String bschluesselKv) {
        this.bschluesselKv = bschluesselKv;
    }

    public String getBschluesselRv() {
        return bschluesselRv;
    }

    public void setBschluesselRv(String bschluesselRv) {
        this.bschluesselRv = bschluesselRv;
    }

    public String getBschluesselAv() {
        return bschluesselAv;
    }

    public void setBschluesselAv(String bschluesselAv) {
        this.bschluesselAv = bschluesselAv;
    }

    public String getBschluesselPv() {
        return bschluesselPv;
    }

    public void setBschluesselPv(String bschluesselPv) {
        this.bschluesselPv = bschluesselPv;
    }

    public String getNameBetrieb() {
        return nameBetrieb;
    }

    public void setNameBetrieb(String nameBetrieb) {
        this.nameBetrieb = nameBetrieb;
    }

    public String getStrassePostfach() {
        return strassePostfach;
    }

    public void setStrassePostfach(String strassePostfach) {
        this.strassePostfach = strassePostfach;
    }

    public String getPlz() {
        return plz;
    }

    public void setPlz(String plz) {
        this.plz = plz;
    }

    public String getOrt() {
        return ort;
    }

    public void setOrt(String ort) {
        this.ort = ort;
    }

    public String getSvTage() {
        return svTage;
    }

    public void setSvTage(String svTage) {
        this.svTage = svTage;
    }

    public String getNatKennzeichen() {
        return natKennzeichen;
    }

    public void setNatKennzeichen(String natKennzeichen) {
        this.natKennzeichen = natKennzeichen;
    }

    public String getMidijobregelung() {
        return midijobregelung;
    }

    public void setMidijobregelung(String midijobregelung) {
        this.midijobregelung = midijobregelung;
    }*/

    @Override
    public String toString() {
        return "MitarbeiterMonat{" +
                /*"nr='" + nr + '\'' +
                ", berater='" + berater + '\'' +
                ", mandant='" + mandant + '\'' +*/
                ", abrechnungsmonat='" + abrechnungsmonat + '\'' +
                ", mitarbeiternummer='" + mitarbeiternummer + '\'' +
                ", nachnameVorname='" + nachnameVorname + '\'' +
                /*", pgr='" + pgr + '\'' +
                ", stkl='" + stkl + '\'' +
                ", antyp='" + antyp + '\'' +
                ", bschluesselKv='" + bschluesselKv + '\'' +
                ", bschluesselRv='" + bschluesselRv + '\'' +
                ", bschluesselAv='" + bschluesselAv + '\'' +
                ", bschluesselPv='" + bschluesselPv + '\'' +*/
                ", svBrutto='" + svBrutto + '\'' +
                /*", nameBetrieb='" + nameBetrieb + '\'' +
                ", strassePostfach='" + strassePostfach + '\'' +
                ", plz='" + plz + '\'' +
                ", ort='" + ort + '\'' +
                ", svTage='" + svTage + '\'' +
                ", natKennzeichen='" + natKennzeichen + '\'' +
                ", midijobregelung='" + midijobregelung + '\'' +*/
                '}';
    }
}
