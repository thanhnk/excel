package my.project.excel.model;

public class Director {
	private String directorId;
	private String country;
	private String companyId;
	private String institutionName;
	private String companyType;
	private String qualification;
	private String qualificationDes;
	private String qualificationDate;
	private String checked;

	@Override
	public String toString() {
		return directorId + "|" + country + "|" + companyId + "|"
				+ institutionName + "|" + companyType + "|" + qualification
				+ "|" + qualificationDes + "|" + qualificationDate + "|"
				+ checked;
	}

	public String getDirectorId() {
		return directorId;
	}

	public void setDirectorId(String directorId) {
		this.directorId = directorId;
	}

	public String getCountry() {
		return country;
	}

	public void setCountry(String country) {
		this.country = country;
	}

	public String getCompanyId() {
		return companyId;
	}

	public void setCompanyId(String companyId) {
		this.companyId = companyId;
	}

	public String getInstitutionName() {
		return institutionName;
	}

	public void setInstitutionName(String institutionName) {
		this.institutionName = institutionName;
	}

	public String getCompanyType() {
		return companyType;
	}

	public void setCompanyType(String companyType) {
		this.companyType = companyType;
	}

	public String getQualification() {
		return qualification;
	}

	public void setQualification(String qualification) {
		this.qualification = qualification;
	}

	public String getQualificationDes() {
		return qualificationDes;
	}

	public void setQualificationDes(String qualificationDes) {
		this.qualificationDes = qualificationDes;
	}

	public String getQualificationDate() {
		return qualificationDate;
	}

	public void setQualificationDate(String qualificationDate) {
		this.qualificationDate = qualificationDate;
	}

	public String getChecked() {
		return checked;
	}

	public void setChecked(String checked) {
		this.checked = checked;
	}
}
