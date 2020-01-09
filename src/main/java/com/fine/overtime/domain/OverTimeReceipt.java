package com.fine.overtime.domain;

import com.fasterxml.jackson.annotation.JsonBackReference;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import javax.persistence.*;

@Entity
@Getter
@Setter
@ToString(exclude = "group")
@Table(name = "overtime_receipt")
public class OverTimeReceipt {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long receipt_id;

    @Column(length = 15)
    private String receipt_date;

    @Column(length = 100)
    private String receipt_name;

    private int receipt_amount;

    @ManyToOne
    @JsonBackReference
    @JoinColumn(name = "group_id")
    private OverTimeGroup group;

}
