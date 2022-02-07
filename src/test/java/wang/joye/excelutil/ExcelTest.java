package wang.joye.excelutil;

import cn.hutool.core.io.FileUtil;
import cn.hutool.json.JSONUtil;
import wang.joye.excelutil.entity.Post;

import java.io.File;
import java.time.LocalDateTime;
import java.util.LinkedList;
import java.util.List;

public class ExcelTest {
    public static void main(String[] args) {

        List<Post> list = new LinkedList<>();
        Post post1 = new Post();
        post1.setAuthor("张三");
        post1.setCoverUrl("https://scpic.chinaz.net/files/pic/pic9/201709/bpic3269.jpg");
        post1.setLevel(1);
        post1.setPublishTime(LocalDateTime.now());
        post1.setTitle("张三的传奇一生");
        list.add(post1);

        Post post2 = new Post();
        post2.setAuthor("张三");
        post2.setCoverUrl("https://th.bing.com/th/id/R.520f0dc0881a866d34cca817ee0e0f56?rik=0tSZUrgZK0FX3A&riu=http%3a%2f%2fimage.hnol.net%2fc%2f2017-12%2f28%2f17%2f20171228174225541-239867.jpg&ehk=HRAJZRx9nyCOmWz7TilqHUc2gpUsBD1rcNlfj8j3ZsM%3d&risl=&pid=ImgRaw&r=0");
        post2.setLevel(2);
        post2.setPublishTime(LocalDateTime.now());
        post2.setTitle("张三的文章1");
        list.add(post2);

        Post post3 = new Post();
        post3.setAuthor("李四");
        post3.setCoverUrl("https://th.bing.com/th/id/OIP.5xwdcGAdwuayJ0eggbN-OQHaE5?pid=ImgDet&rs=1");
        post3.setLevel(1);
        post3.setPublishTime(LocalDateTime.now());
        post3.setTitle("李四的文章2");
        list.add(post3);

        // 使用hutool工具类, 从文件中读取配置
        String configStr = FileUtil.readUtf8String("export_config.json");
        // 将json字符串转为list
        List<BeanToExcelConfig> configs = JSONUtil.toList(configStr, BeanToExcelConfig.class);

        // 初始化配置
        ExcelGenerateUtil.initConfig(configs);
        // 生成excel文件
        File file = ExcelGenerateUtil.generateFile(list, "post", "测试标题");
        System.out.println(file.getAbsolutePath());
    }
}
