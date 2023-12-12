import configparser
import os
import sqlalchemy as sa
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import Integer, Column, String, Date, Numeric
from pathlib import Path


#Import Config Settings
Config = configparser.ConfigParser()

#CONFIG
working_directory = os.getcwd()
Config.read('./settings.ini')

class DatabaseConnection:
    """Connect to SQL Server Domo Database"""
    _db_connection = None
    _db_cur = None

    def __init__(self):
        self.connectionstring = "mssql+pyodbc://" + Config.get('DomoDB','username') + ":" \
                                + Config.get('DomoDB','password') + "@"\
                                + Config.get('DomoDB','server') + ":" \
                                + Config.get('DomoDB','port') + "/"\
                                + Config.get('DomoDB','database') \
                                + "?Encrypt=yes&TrustServerCertificate=yes" + "&" \
                                + Config.get('DomoDB', 'driver')


        self.engine = sa.create_engine(self.connectionstring)
        self.connection = self.engine.connect()


# define declarative base
Base = declarative_base()

# create an engine
db = DatabaseConnection()

# reflect current database engine to metadata
# metadata = sa.MetaData(db.engine)
metadata = sa.MetaData(db.engine)
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
    baseline_diagnostic = Column("BaselineDiagnostic(Y/N)", String)
    most_recent_diagnostic = Column("MostRecentDiagnostic(Y/N)", String)
    duration = Column("Duration(min)", Integer)
    rush_flag = Column("RushFlag", String)
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
    reading_comprehension_overall_scale_score = Column("ReadingComprehension:OverallScaleScore", Integer)
    reading_comprehension_overall_placement = Column("ReadingComprehension:OverallPlacement", String)
    reading_comprehension_overall_relative_placement = Column("ReadingComprehension:OverallRelativePlacement", String)
    reading_comprehension_literature_scale_score = Column("ReadingComprehension:LiteratureScaleScore", Integer)
    reading_comprehension_literature_placement = Column("ReadingComprehension:LiteraturePlacement", String)
    reading_comprehension_literature_relative_placement = Column("ReadingComprehension:LiteratureRelativePlacement", String)
    reading_comprehension_informational_text_scale_score = Column("ReadingComprehension:InformationalTextScaleScore", Integer)
    reading_comprehension_informational_text_placement = Column("ReadingComprehension:InformationalTextPlacement", String)
    reading_comprehension_informational_text_relative_placement = Column("ReadingComprehension:InformationalTextRelativePlacement", String)
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
    baseline_diagnostic = Column("BaselineDiagnostic(Y/N)", String)
    most_recent_diagnostic = Column("MostRecentDiagnostic(Y/N)", String)
    duration = Column("Duration(min)", Integer)
    rush_flag = Column("RushFlag", String)
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
    school_name = Column("SchoolName", String)
    last_name = Column("LastName", String)
    first_name = Column("FirstName", String)
    middle_name = Column("MiddleName", String)
    grade = Column("Grade", String)
    student_id = Column("StudentId", String, primary_key=True,autoincrement=False)
    iready_student_id = Column("iReadyStudentId", String)
    gender = Column("Gender", String)
    date_of_birth = Column("DateOfBirth", Date)
    ethnicity = Column("Ethnicity",String)
    esl = Column("ESL", String)
    SchoolNPSIS = Column("SchoolNPSIS", String)
    NPSCode = Column("NPSCode", String)
    referral_status = Column("ReferralStatus", String)
    consent_status = Column("ConsentStatus", String)
    status = Column("Status", String) 
 

class LoadProduction:
    """Load AI Production Tables"""
   
    def load_production_tables(self):

        db = DatabaseConnection()
        
        trans = db.connection.begin()
        db.connection.execute("AI.AILoadTables")
        trans.commit()

