package com.eitech1.chartv.Entity;

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

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@Entity
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class Tab {
	@Id
	@GeneratedValue
	@Column(name = "tab_id")
	private int tabId;
	
	@Column(name = "tab_name")
	private String tabName;
	
	@Column(name = "tab_topic")
	private String tabTopic;
	
	@OneToMany(mappedBy = "tab", cascade = CascadeType.ALL)
	private List<DataSet> data;
	
	@ManyToOne(fetch = FetchType.LAZY)
	@JoinColumn(name = "sheet_id")
	private SheetEx sheet;

}
