package top.simba1949.dao;

import top.simba1949.domain.User;

import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * @author anthony
 * @date 2021/3/14 19:09
 */
public interface UserDao {
    /**
     * 获取数据
     * @return
     */
    static List<User> getData(){
        return Stream.iterate(0, n -> n + 1)
                .limit(10)
                .map(index -> User.builder()
                        .id((long) index).username("username:" + index)
                        .height(BigDecimal.valueOf(index).multiply(new BigDecimal("1.011")).setScale(2, BigDecimal.ROUND_HALF_UP))
                        .birthday(LocalDateTime.now()).build()
                ).collect(Collectors.toList());
    }

    /**
     * 获取标题
     * @return
     */
    static List<String> getHeader(){
        return Arrays.asList("序号", "用户名", "身高", "生日");
    }
}
