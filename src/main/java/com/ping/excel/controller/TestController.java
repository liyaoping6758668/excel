package com.ping.excel.controller;


import com.ping.excel.pojo.User;
import com.ping.excel.utils.ExcelExport2;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import javax.servlet.http.HttpServletResponse;
import java.util.ArrayList;

/**
 * @Author：huang
 * @Date：2019-09-21 13:13
 * @Description：<描述>
 */
@Controller
public class TestController {

    @RequestMapping("/test")
    public void testExprotExcel(HttpServletResponse response){

        //创建一个数组用于设置表头
        String[] arr = new String[]{"ID","用户名","密码","备注"};

        ArrayList arrayList=new ArrayList();
        for (int i = 0; i <5 ; i++) {
            User user=new User();
            user.setUid(i+"");
            user.setUsername("liyaoping"+i);
            user.setPassword("ping"+i);
            user.setRemark("666");
            arrayList.add(user);
        }

        //调用Excel导出工具类
        ExcelExport2.export(response,arrayList,arr);

    }

}

