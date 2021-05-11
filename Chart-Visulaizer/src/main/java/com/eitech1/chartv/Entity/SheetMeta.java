package com.eitech1.chartv.Entity;

import java.util.List;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.Id;
import javax.persistence.OneToMany;

import lombok.Data;

@Data
@Entity
public class SheetMeta {
	
	@Id
	private int id;
	
	@Column(name = "sheet_name")
	private String sheetName;
	
	@Column(name = "sheet_code")
	private String sheetCode;
	
	@OneToMany(mappedBy = "sheetMeta")
	private List<SheetEx> sheets;
	

}
