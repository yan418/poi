package com.im.modules.dao;

import com.im.modules.entities.User;

import java.util.List;

public interface UserDao {

    List<User> getUser();

    int addUser(User user);

}
