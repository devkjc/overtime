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
@Table(name = "overtime_people")
public class OverTimePeople {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long people_id;

    private String people_position;

    private String people_name;

    private String people_duties;

    @ManyToOne
    @JsonBackReference
    @JoinColumn(name = "group_id")
    private OverTimeGroup group;

}
