package pub.developers.docautogenbyexcel.repository;

import org.springframework.data.domain.Page;
import org.springframework.data.domain.Pageable;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;
import pub.developers.docautogenbyexcel.entity.DocumentEntity;

import java.time.LocalDateTime;
import java.util.List;
import java.util.Optional;

/**
 * 文档Repository接口
 */
@Repository
public interface DocumentRepository extends JpaRepository<DocumentEntity, Long> {

    /**
     * 根据输出ID查找文档
     */
    Optional<DocumentEntity> findByOutputId(String outputId);

    /**
     * 根据文件名查找文档
     */
    Optional<DocumentEntity> findByOutputFileName(String outputFileName);

    /**
     * 查找所有文档，按创建时间倒序
     */
    List<DocumentEntity> findAllByOrderByCreatedAtDesc();

    /**
     * 分页查询文档，按创建时间倒序
     */
    Page<DocumentEntity> findAllByOrderByCreatedAtDesc(Pageable pageable);

    /**
     * 根据创建时间范围查找文档
     */
    List<DocumentEntity> findByCreatedAtBetween(LocalDateTime start, LocalDateTime end);

    /**
     * 根据存储类型查找文档
     */
    List<DocumentEntity> findByStorageTypeOrderByCreatedAtDesc(String storageType);

    /**
     * 统计文档数量
     */
    @Query("SELECT COUNT(d) FROM DocumentEntity d")
    long countAll();

    /**
     * 统计指定时间范围内的文档数量
     */
    @Query("SELECT COUNT(d) FROM DocumentEntity d WHERE d.createdAt BETWEEN ?1 AND ?2")
    long countByCreatedAtBetween(LocalDateTime start, LocalDateTime end);

    /**
     * 删除指定时间之前的文档
     */
    void deleteByCreatedAtBefore(LocalDateTime cutoffTime);
}

