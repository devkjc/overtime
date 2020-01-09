package com.fine.overtime.repo;

import com.fine.overtime.domain.OverTimeGroup;
import com.fine.overtime.domain.OverTimePeople;
import com.fine.overtime.domain.OverTimeReceipt;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.querydsl.QuerydslPredicateExecutor;

import java.util.List;

public interface GroupRepo extends JpaRepository<OverTimeGroup, Long>, QuerydslPredicateExecutor<OverTimeGroup> {

}
