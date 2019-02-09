package com.dangelis.entity;

import java.util.List;

/**
 * Created by Altran on 18/04/17.
 */
public class Room {
    private String _name;
    private String _email;
    private List<Appointment> list;

    public Room(String _name, String _email, List<Appointment> list) {
        this._name = _name;
        this._email = _email;
        this.list = list;
    }

    public String get_name() {
        return _name;
    }

    public void set_name(String _name) {
        this._name = _name;
    }

    public String get_email() {
        return _email;
    }

    public void set_email(String _email) {
        this._email = _email;
    }

    public List<Appointment> getList() {
        return list;
    }

    public void setList(List<Appointment> list) {
        this.list = list;
    }
}
