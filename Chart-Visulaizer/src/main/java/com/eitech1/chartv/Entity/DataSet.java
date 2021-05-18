package com.eitech1.chartv.Entity;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.FetchType;
import javax.persistence.GeneratedValue;
import javax.persistence.Id;
import javax.persistence.JoinColumn;
import javax.persistence.Lob;
import javax.persistence.ManyToOne;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;



@Data
@Entity
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class DataSet {
	@Id
	@GeneratedValue
	@Column(name = "data_set_id")
	private int dataSetId;
	
	@Column(name = "json_data")
	@Lob
	private String jsonData;
	
	@ManyToOne(fetch = FetchType.LAZY)
	@JoinColumn(name = "tab_id")
	private Tab tab;

}
