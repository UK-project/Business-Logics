package com.eitech1.chartv.respository;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import com.eitech1.chartv.Entity.SheetEx;

@Repository
public interface SheetExRepository extends JpaRepository<SheetEx, Integer> {

}
