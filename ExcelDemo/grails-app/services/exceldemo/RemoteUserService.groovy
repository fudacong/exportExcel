package exceldemo

import grails.transaction.Transactional

@Transactional
class RemoteUserService {

    /*假设这里是从数据库查出的数据*/
    List<UserBean> findUserList(){
        List<UserBean> list = new ArrayList<UserBean>()
        for(int i = 0 ; i < 30 ; i ++){
            UserBean userBean = new UserBean();
            userBean.userName = "小明${i}"
            userBean.age = i;
            userBean.company = "蜀国${i}"
            userBean.hobby = "做作业"
            userBean.password = "1232323${i}"
            userBean.sex = "男"
            list.add(userBean)
        }
        return  list;
    }
}
