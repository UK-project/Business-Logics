package com.eitech1.chartv.Entity;

import java.util.Date;
import java.util.List;

import javax.persistence.CascadeType;
import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.FetchType;
import javax.persistence.GeneratedValue;
import javax.persistence.Id;
import javax.persistence.JoinColumn;
import javax.persistence.ManyToOne;
import javax.persistence.OneToMany;

import lombok.Builder;
import lombok.Data;

@Data
@Entity
@Builder
public class SheetEx {
	@Id
	@GeneratedValue
	@Column(name = "sheet_id")
	private int sheetId;
	
	@Column(name = "uploaded_date")
	private Date uploadedDate;
	
	@OneToMany(mappedBy = "sheet", cascade = CascadeType.ALL)
	private List<Tab> tabs;
	
	@ManyToOne(fetch = FetchType.LAZY)
	@JoinColumn(name = "sheet_meta_id")
	private SheetMeta sheetMeta;
	

}
