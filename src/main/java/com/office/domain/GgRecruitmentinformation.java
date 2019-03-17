package com.office.domain;

import lombok.Data;

import java.util.Date;


@Data
public class GgRecruitmentinformation  {

	private Long id;

	/**
	 * 公司名称
	 */
	private String companyname;
	private String companyaddress;

	/**
	 * 发布时间
	 */
	private Date publishdate;

	/**
	 * 岗位
	 */
	private String post;

	/**
	 * 公司介绍
	 */
	private String companyintroduce;

	/**
	 * 工资
	 */
	private String salary;

	/**
	 * 岗位职责
	 */
	private String duty;

	/**
	 * 任职资格
	 */
	private String requirement;

	/**
	 * 学历
	 */
	private String education;

	/**
	 * 招聘人数
	 */
	private String count;

	/**
	 * 工作经验
	 */
	private String employmenttime;

	/**
	 * 公司定位ip
	 */
	private String companyip;

	/**
	 * 操作员id
	 */
	private Long operaterid;

	/**
	 *   本地   外地
	 */
	private String location;

	/**
	 * 工种
	 */
	private String worktype;

	/**
	 * 要求专业种类
	 */
	private String major;

	/**
	 * 工作性质：全职/兼职
	 */
	private String type;

	/**
	 * 待遇详情
	 */
	private String treatment;

	/**
	 * 联系人
	 */
	private String linkman;

	/**
	 * 电话
	 */
	private String phone;

	/**
	 * 0代表上架 1代表下架
	 */
	private String status;

	/**
	 * 工作地址
	 */
	private String workplace;

	/**
	 * 所属单位端公司id
	 */
	private Long unitendid;
	
}

