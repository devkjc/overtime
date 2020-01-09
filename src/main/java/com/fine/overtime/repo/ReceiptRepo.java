package com.fine.overtime.repo;

import com.fine.overtime.domain.OverTimeGroup;
import com.fine.overtime.domain.OverTimeReceipt;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Modifying;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.querydsl.QuerydslPredicateExecutor;
import org.springframework.transaction.annotation.Transactional;

import java.util.ArrayList;

public interface ReceiptRepo extends JpaRepository<OverTimeReceipt, Long>, QuerydslPredicateExecutor<OverTimeReceipt> {

    @Query
    ArrayList<OverTimeReceipt> findByGroup(OverTimeGroup group);

    @Modifying
    @Transactional
    @Query("delete from OverTimeReceipt r where r.group.groupId = :group_id")
    void deleteByGroup(Long group_id);

}
