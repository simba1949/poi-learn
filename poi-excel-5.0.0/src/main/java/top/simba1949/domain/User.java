package top.simba1949.domain;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.io.Serializable;
import java.math.BigDecimal;
import java.time.LocalDateTime;

/**
 * @author anthony
 * @date 2021/3/14 18:57
 */
@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class User implements Serializable {
    private static final long serialVersionUID = -3064936650154013381L;

    private Long id;
    private String username;
    private BigDecimal height;
    private LocalDateTime birthday;
}
