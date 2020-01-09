package com.fine.overtime.domain;

import com.fasterxml.jackson.annotation.JsonManagedReference;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;
import org.hibernate.annotations.CreationTimestamp;

import javax.persistence.*;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.List;

@Entity
@Getter
@Setter
@ToString(exclude = {"peopleList","receiptList"})
@Table(name = "overtime_group")
public class OverTimeGroup {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    @Column(name = "group_id")
    private Long groupId;

    private String group_name;

    private String group_officer_dept;

    private String group_officer_name;

    private String group_support;

    private String group_period;

    private String group_subject_name;

    private String group_place;

    @OneToMany(fetch = FetchType.LAZY, cascade = CascadeType.ALL, orphanRemoval = true)
    @JoinColumn(name = "group_id")
    @JsonManagedReference
    private List<OverTimePeople> peopleList;

    @OneToMany(fetch = FetchType.LAZY, cascade = CascadeType.ALL, orphanRemoval = true)
    @JoinColumn(name = "group_id")
    @JsonManagedReference
    private List<OverTimeReceipt> receiptList;

    @CreationTimestamp
    @Column(updatable = false)
    private Timestamp group_createTime;
}
