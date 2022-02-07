package wang.joye.excelutil.entity;

import lombok.Data;

import java.time.LocalDateTime;

/**
 * 文章类
 */
@Data
public class Post {
    /**
     * 作者
     */
    private String author;
    /**
     * 标题
     */
    private String title;
    /**
     * 封面url
     */
    private String coverUrl;
    /**
     * 等级
     */
    private Integer level;
    /**
     * 发表时间
     */
    private LocalDateTime publishTime;
}
