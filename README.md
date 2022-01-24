### 说明

本项目用于excel生成与导出，依赖poi和hutool 可以写出到文件，可以导出到流(如HttpServletResponse)
The project is an excel util for java, which can generates excel by some config, also supports export to stream.

### 添加依赖

```xml

<dependency>
    <groupId>wang.joye</groupId>
    <artifactId>excel-util</artifactId>
    <version>${latest.version}</version>
</dependency>
```

注意，其中的latest.version为最新版本

### 使用示例

```java
import lombok.Data;

import java.time.LocalDateTime;

/**
 * 文章类
 */
@Data
public class Post {
    private String author;
    private Sting coverBase64;
    private Integer level;
    private LocalDateTime publishTime;
    private Sting title;
}
```

```json
[
  {
    "name": "post",
    "remark": "这里是备注",
    "fields": [
      {
        "name": "author",
        "type": "string",
        "title": "作者",
        "mergeRepeatRow": "true"
      },
      {
        "name": "coverBase64",
        "type": "base64",
        "title": "封面"
      },
      {
        "name": "level",
        "type": "map",
        "title": "等级",
        "valueMap": {
          "1": "一级",
          "2": "二级"
        }
      },
      {
        "name": "publishTime",
        "type": "datetime",
        "title": "发表时间"
      },
      {
        "name": "title",
        "type": "string",
        "title": "标题"
      }
    ]
  }
]
```

```
// 生成excel文件
ExcelGenerateUtil.generate2File(list, "post");

HttpServletResponse response;
OutputStream out = response.getOutputStream();
List<Post> list = new LinkedList<>();
// 导出到流
ExcelGenerateUtil.generateAndWrite(list, "post", out);
```