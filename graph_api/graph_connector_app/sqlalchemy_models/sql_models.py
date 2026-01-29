import configparser
import os
import sqlalchemy as sa
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import Integer, Column, String, Date, DateTime, Numeric, Boolean
from pathlib import Path


#Import Config Settings
Config = configparser.ConfigParser()

#CONFIG
working_directory = os.getcwd()
Config.read('./settings.ini')

class DatabaseConnection:
    """Connect to LEARN Behavioral SQL Server Database"""
    _db_connection = None
    _db_cur = None

    def __init__(self, section: str = "DomoDB"):
        cfg = Config[section]
        self.connectionstring = (
            f"mssql+pyodbc://{cfg['username']}:{cfg['password']}@"
            f"{cfg['server']}:{cfg['port']}/{cfg['database']}"
            f"?Encrypt=yes&TrustServerCertificate=yes&{cfg['driver']}"
        )
        # Enable fast_executemany to speed up bulk inserts (pyodbc-specific optimization)
        self.engine = sa.create_engine(
            self.connectionstring,
            fast_executemany=True,
        )
        self.connection = self.engine.connect()


# define declarative base
Base = declarative_base()

# create DomoIntegration engine
db = DatabaseConnection("DomoDB")

# reflect current database engine to metadata
# metadata = sa.MetaData(db.engine)
metadata = sa.MetaData(db.engine)
metadata.reflect()

# create Integration engine
db_integration = DatabaseConnection("Integration")

# reflect current database engine to metadata
# metadata = sa.MetaData(db_integration.engine)
metadata = sa.MetaData(db_integration.engine)
metadata.reflect()



# build your ReadingiReady class on existing `AI.ReadingiReadyStaging` table
class ReadingiReady(Base):
    """Reading iReady Model"""
    __tablename__ = "ReadingIreadyStaging"
    __table_args__ = {"schema": "AI"}
    #column_not_exist_in_db = Column(Integer, primary_key=True,autoincrement=False) # just add for sake of this error, dont add in db
    district = Column("District", String)
    subject = Column("Subject", String)
    last_name = Column("LastName", String)
    first_name = Column("FirstName", String)
    student_id = Column("StudentID", String, primary_key=True,autoincrement=False)
    student_grade = Column("StudentGrade", String)
    academic_year = Column("AcademicYear", String)
    school = Column("School", String)
    enrolled = Column("Enrolled", String)
    district_state_id = Column("DistrictStateID", String)
    account_state_id = Column("AccountStateID", String)
    school_state_id = Column("SchoolStateID", String)
    student_state_id = Column("StudentStateID", String)
    user_name = Column("UserName", String)
    sex = Column("Sex", String)
    hispanic_or_latino = Column("HispanicorLatino", String)
    race = Column("Race", String)
    english_language_learner = Column("EnglishLanguageLearner", String)
    special_education = Column("SpecialEducation", String)
    economic_disadvantaged = Column ("EconomicallyDisadvantaged", String)
    migrant = Column("Migrant", String)
    classes = Column("Class(es)", String)
    class_teachers = Column("ClassTeacher(s)", String)
    report_groups = Column("ReportGroup(s)", String)
    start_date = Column("StartDate", Date)
    completion_date = Column("CompletionDate", Date)
    norming_window = Column("NormingWindow", String)
    baseline_diagnostic = Column("BaselineDiagnostic(Y/N)", String)
    most_recent_diagnostic = Column("MostRecentDiagnosticYTD(Y/N)", String)
    duration = Column("Duration(min)", Integer)
    rush_flag = Column("RushFlag", String)
    read_aloud = Column("ReadAloud", String)
    overall_scale_score = Column("OverallScaleScore", Integer)
    overall_placement = Column("OverallPlacement", String)
    overall_relative_placement = Column("OverallRelativePlacement", String)
    percentile = Column("Percentile", Integer)
    grouping = Column("Grouping", Integer)
    lexile_measure = Column("LexileMeasure", String)
    lexile_range = Column("LexileRange", String)
    phonological_awareness_scale_score = Column("PhonologicalAwarenessScaleScore", Integer)
    phonological_awareness_placement = Column("PhonologicalAwarenessPlacement", String)
    phonological_awareness_relative_placement = Column("PhonologicalAwarenessRelativePlacement", String)
    phonics_scale_score = Column("PhonicsScaleScore", Integer)
    phonics_placement = Column("PhonicsPlacement", String)
    phonics_relative_placement = Column("PhonicsRelativePlacement", String)
    high_frequency_words_scale_score = Column("High-FrequencyWordsScaleScore", Integer)
    high_frequency_words_placement = Column("High-FrequencyWordsPlacement", String)
    high_frequency_words_relative_placement = Column("High-FrequencyWordsRelativePlacement", String)
    vocabulary_scale_score = Column("VocabularyScaleScore", Integer)
    vocabulary_placement = Column("VocabularyPlacement", String)
    vocabulary_relative_placement = Column("VocabularyRelativePlacement", String)
    comprehension_overall_scale_score = Column("Comprehension:OverallScaleScore", Integer)
    comprehension_overall_placement = Column("Comprehension:OverallPlacement", String)
    comprehension_overall_relative_placement = Column("Comprehension:OverallRelativePlacement", String)
    comprehension_literature_scale_score = Column("Comprehension:LiteratureScaleScore", Integer)
    comprehension_literature_placement = Column("Comprehension:LiteraturePlacement", String)
    comprehension_literature_relative_placement = Column("Comprehension:LiteratureRelativePlacement", String)
    comprehension_informational_text_scale_score = Column("Comprehension:InformationalTextScaleScore", Integer)
    comprehension_informational_text_placement = Column("Comprehension:InformationalTextPlacement", String)
    comprehension_informational_text_relative_placement = Column("Comprehension:InformationalTextRelativePlacement", String)
    diagnostic_gain = Column("DiagnosticGain", Numeric)
    annual_typical_growth_measure = Column("AnnualTypicalGrowthMeasure", Integer)
    annual_stretch_growth_measure = Column("AnnualStretchGrowthMeasure", Integer)
    percent_progresstoAnnual_typical_growth = Column("PercentProgresstoAnnualTypicalGrowth(%)", Numeric)
    percent_progresstoAnnual_stretch_growth = Column("PercentProgresstoAnnualStretchGrowth(%)", Numeric)
    mid_on_grade_level_scale_score = Column("MidOnGradeLevelScaleScore", Integer)
    reading_difficulty_indicator = Column("ReadingDifficultyIndicator(Y/N)", String)


# build your MathiReady class on existing `AI.MathiReadyStaging` table
class MathiReady(Base):
    """Math iReady Model"""
    __tablename__ = "MathIreadyStaging"
    __table_args__ = {"schema": "AI"}
    #column_not_exist_in_db = Column(Integer, primary_key=True,autoincrement=False) # just add for sake of this error, dont add in db
    district = Column("District", String)
    subject = Column("Subject", String)
    last_name = Column("LastName", String)
    first_name = Column("FirstName", String)
    student_id = Column("StudentID", String, primary_key=True,autoincrement=False)
    student_grade = Column("StudentGrade", String)
    academic_year = Column("AcademicYear", String)
    school = Column("School", String)
    enrolled = Column("Enrolled", String)
    district_state_id = Column("DistrictStateID", String)
    account_state_id = Column("AccountStateID", String)
    school_state_id = Column("SchoolStateID", String)
    student_state_id = Column("StudentStateID", String)
    user_name = Column("UserName", String)
    sex = Column("Sex", String)
    hispanic_or_latino = Column("HispanicorLatino", String)
    race = Column("Race", String)
    english_language_learner = Column("EnglishLanguageLearner", String)
    special_education = Column("SpecialEducation", String)
    economic_disadvantaged = Column("EconomicallyDisadvantaged", String)
    migrant = Column("Migrant", String)
    classes = Column("Class(es)", String)
    class_teachers = Column("ClassTeacher(s)", String)
    report_groups = Column("ReportGroup(s)", String)
    start_date = Column("StartDate", Date)
    completion_date = Column("CompletionDate", Date)
    norming_window = Column("NormingWindow", String)
    baseline_diagnostic = Column("BaselineDiagnostic(Y/N)", String)
    most_recent_diagnostic = Column("MostRecentDiagnosticYTD(Y/N)", String)
    duration = Column("Duration(min)", Integer)
    rush_flag = Column("RushFlag", String)
    read_aloud = Column("ReadAloud", String)
    overall_scale_score = Column("OverallScaleScore", Integer)
    overall_placement = Column("OverallPlacement", String)
    overall_relative_placement = Column("OverallRelativePlacement", String)
    percentile = Column("Percentile", Integer)
    grouping = Column("Grouping", Integer)
    quantile_measure = Column("QuantileMeasure", String)
    quantile_range = Column("QuantileRange", String)
    number_and_operations_scale_score = Column("NumberandOperationsScaleScore", Integer)
    number_and_operations_placement = Column("NumberandOperationsPlacement", String)
    number_and_operations_relative_placement = Column("NumberandOperationsRelativePlacement", String)
    algebra_and_algebraic_thinking_scale_score = Column("AlgebraandAlgebraicThinkingScaleScore", Integer)
    algebra_and_algebraic_thinking_placement = Column("AlgebraandAlgebraicThinkingPlacement", String)
    algebra_and_algebraic_thinking_relative_placement = Column("AlgebraandAlgebraicThinkingRelativePlacement", String)
    measurement_and_data_scale_score = Column("MeasurementandDataScaleScore", Integer)
    measurement_and_data_placement = Column("MeasurementandDataPlacement", String)
    measurement_and_data_relative_placement = Column("MeasurementandDataRelativePlacement", String)
    geometry_scale_score = Column("GeometryScaleScore", Integer)
    geometry_placement = Column("GeometryPlacement", String)
    geometry_relative_placement = Column("GeometryRelativePlacement", String)
    diagnostic_language = Column("DiagnosticLanguage", String)
    diagnostic_gain = Column("DiagnosticGain", Numeric)
    annual_typical_growth_measure = Column("AnnualTypicalGrowthMeasure", Integer)
    annual_stretch_growth_measure = Column("AnnualStretchGrowthMeasure", Integer)
    percent_progresstoAnnual_typical_growth = Column("PercentProgresstoAnnualTypicalGrowth(%)", Numeric)
    percent_progresstoAnnual_stretch_growth = Column("PercentProgresstoAnnualStretchGrowth(%)", Numeric)
    mid_on_grade_level_scale_score = Column("MidOnGradeLevelScaleScore", Integer)

# build your Eligibility class on existing `AI.EligibiltyStaging` table
class Eligibility(Base):
    """Eligibility Model"""
    __tablename__ = "EligibilityStaging"
    __table_args__ = {"schema": "AI"}
    district = Column("District", String)
    subject = Column("Subject", String)
    academic_year = Column("AcademicYear", String)
    school_name = Column("SchoolName", String)
    project_code = Column("ProjectCode", String)
    school_code = Column("SchoolCode", String)
    last_name = Column("StudentLastName", String)
    first_name = Column("StudentFirstName", String)
    middle_name = Column("StudentMiddleName", String)
    grade = Column("GradeLevel", String)
    referral_type = Column("ReferralType",String)
    student_id = Column("StudentId", String, primary_key=True,autoincrement=False)
    iready_student_id = Column("i-ReadyStudentUserId", String)
    gender = Column("Gender", String)
    date_of_birth = Column("DateOfBirth", Date)
    ethnicity = Column("Ethnicity",String)
    esl = Column("ESL", String)
    school_npsis = Column("SchoolNPSIS", String)
    nps_code = Column("NPSCode", String)
    counseling = Column("Counseling", String)
    participates_in_title_i = Column("ParticipatesinTitleI", String)

    #consent_status = Column("ConsentStatus", String)
    #status = Column("Status", String) 

# build your UserInformation class on existing `IMS.UserInformationListStaging` table
class UserInformation(Base):
    """User Information Model"""
    __tablename__ = "UserInformationListStaging"
    __table_args__ = {"schema": "IMS"}
    content_type_id = Column("ContentTypeID", String)
    name = Column("Name", String)
    compliance_asset_id = Column("ComplianceAssetId", String)
    account = Column("Account", String)
    email = Column("EMail", String)
    other_mail = Column("OtherMail", String)
    user_expiration = Column("UserExpiration", DateTime)
    user_last_deletion_time = Column("UserLastDeletionTime", DateTime)
    mobile_number = Column("MobileNumber", String)
    about_me = Column("AboutMe", String)
    sip_address = Column("SIPAddress", String)
    is_site_admin = Column("IsSiteAdmin", Boolean)
    deleted = Column("Deleted", Boolean)
    hidden = Column("Hidden", Boolean)
    picture = Column("Picture", String)
    department = Column("Department", String)
    job_title = Column("JobTitle", String)
    first_name = Column("FirstName", String)
    last_name = Column("LastName", String)
    work_phone = Column("WorkPhone", String)
    user_name = Column("UserName", String)
    website = Column("WebSite", String)
    ask_me_about = Column("AskMeAbout", String)
    office = Column("Office", String)
    modified = Column("Modified", DateTime)
    id = Column("Id", Integer, primary_key=True, autoincrement=False)
    content_type = Column("ContentType", String)
    created = Column("Created", DateTime)
    created_by_id = Column("CreatedById", Integer)
    modified_by_id = Column("ModifiedById", Integer)
    owshiddenversion = Column("Owshiddenversion", String)
    version = Column("Version", String)
    path = Column("Path", String)



# build your AssetManagement class on existing `IMS.AssetManagementListStaging` table
class AssetManagement(Base):
    """Asset Management Model"""
    __tablename__ = "AssetManagementListStaging"
    __table_args__ = {"schema": "IMS"}
    id = Column("Id", Integer, primary_key=True, autoincrement=False)
    content_type_id = Column("ContentTypeID", String)
    content_type = Column("ContentType", String)
    title = Column("Title", String)
    modified = Column("Modified", DateTime)
    created = Column("Created", DateTime)
    created_by_id = Column("CreatedById", Integer)
    modified_by_id = Column("ModifiedById", Integer)
    owshiddenversion = Column("Owshiddenversion", String)
    version = Column("Version", String)
    path = Column("Path", String)
    compliance_asset_id = Column("ComplianceAssetId", String)
    status_value = Column("StatusValue", String)
    manufacturer = Column("Manufacturer", String)
    model = Column("Model", String)
    color_value = Column("ColorValue", String)
    tablet_number = Column("TabletNumber", String)
    current_owner_id = Column("CurrentOwnerId", Integer)
    previous_owner_id = Column("PreviousOwnerId", Integer)
    due_date = Column("DueDate", DateTime)
    current_owner_previous_owner_id = Column("CurrentOwnerPreviousOwnerId", Integer)
    date_assigned = Column("DateAssigned", DateTime)
    date_reached_out_for_collection = Column("DateReachedOutForCollection", DateTime)
    staff_last_assigned_to_email_id = Column("StaffLastAssignedToEmailId", Integer)
    assigned_by_id = Column("AssignedById", Integer)
    location_value = Column("LocationValue", String)
    tracking_number = Column("TrackingNumber", String)
    has_a_working_charger = Column("HasAWorkingCharger", String)
    staff_last_assigned_to_full_name = Column("StaffLastAssignedToFullName", String)
    color_tag = Column("ColorTag", String)
    activity_count = Column("ActivityCount", String)
    current_owner_name_txt = Column("CurrentOwnerName_txt", String)
    current_owner_email_txt = Column("CurrentOwnerEMail_txt", String)
    assigned_by_name_txt = Column("AssignedByName_txt", String)
    assigned_by_email_txt = Column("AssignedByEmail_txt", String)

 



class LoadProduction:
    """Load AI Production Tables"""
   
    def load_production_tables(self,option=3):

        db = DatabaseConnection()
        
        trans = db.connection.begin()
        db.connection.execute(f"EXEC AI.AILoadTables @Option = {option}")
        trans.commit()

