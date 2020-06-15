package com.fine.overtime.repo;

import com.fine.overtime.domain.OverTimeGroup;
import com.fine.overtime.domain.OverTimePeople;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Modifying;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.querydsl.QuerydslPredicateExecutor;
import org.springframework.transaction.annotation.Transactional;

import java.util.ArrayList;

public interface PeopleRepo extends JpaRepository<OverTimePeople, Long>, QuerydslPredicateExecutor<OverTimePeople> {

    @Query
    ArrayList<OverTimePeople> findByGroup(OverTimeGroup group);

    @Query
    @Modifying
    @Transactional
    void deleteByGroupIsNull();

}
