package com.eitech1.chartv.respository;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import com.eitech1.chartv.Entity.Tab;

@Repository
public interface TabRepository extends JpaRepository<Tab, Integer>{

}
