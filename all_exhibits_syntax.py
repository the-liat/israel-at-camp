exhibit_syntax ={
    1:"""MEANS TABLES=Age_continuous BY Campers_vs_staff
            /CELLS=MEAN COUNT STDDEV.
        CROSSTABS
            /TABLES=Campers_vs_staff BY Gender_2cat Denomination_RC Attended_Jewish_day_school_RC
            Attended_overnight_camp_RC Attended_supp_school_RC Attended_Jewish_youth_group_RC
            How_often_attended_religious_service Times_visited_Israel_new_RC
            /FORMAT=AVALUE TABLES
            /CELLS=COUNT ROW
            /COUNT ROUND CELL.
        """,
    2:"""CROSSTABS
            /TABLES=Campers_vs_staff BY How_many_Israelis_at_camp
            /FORMAT=AVALUE TABLES
            /STATISTICS=CHISQ
            /CELLS=COUNT ROW
            /COUNT ROUND CELL.""",
    3:"""CROSSTABS
            /TABLES=Campers_vs_staff BY How_often_hear_Hebrew_spoken_at_camp
            /FORMAT=AVALUE TABLES
            /STATISTICS=CHISQ
            /CELLS=COUNT ROW
            /COUNT ROUND CELL.
            """,
    4:"""MULT RESPONSE GROUPS=$Feel_Israel (Feel_Israel_at_camp_Israel_Day
            Feel_Israel_at_camp_counsellors_and_Staff Feel_Israel_at_camp_mifkad_flag_pole
            Feel_Israel_at_camp_dining_hall_chedar_ochel Feel_Israel_at_camp_Shabbat_services_Havdallah
            Feel_Israel_at_camp_words_on_buildings Feel_Israel_at_camp_Israel_specialty
            Feel_Israel_at_camp_night_activity_peula Feel_Israel_at_camp_announcement (1))
            /VARIABLES=Campers_vs_staff(1 2)
            /TABLES=$Feel_Israel BY Campers_vs_staff
            /CELLS=COLUMN
            /BASE=CASES.""",
    5:"""CROSSTABS
            /TABLES=Campers_vs_staff BY How_often_have_Israel_programming_at_camp
            /FORMAT=AVALUE TABLES
            /STATISTICS=CHISQ
            /CELLS=COUNT ROW
            /COUNT ROUND CELL.
            """,
    6:"""MULT RESPONSE GROUPS=$Content (Israel_programs_history Israel_programs_music Israel_programs_people
            Israel_programs_army Israel_programs_cooking Israel_programs_dancing Israel_programs_war
            Israel_programs_technology_and_inventions (1))
            /VARIABLES=Campers_vs_staff(1 2)
            /TABLES=$Content BY Campers_vs_staff
            /CELLS=COLUMN
            /BASE=CASES.
            """,
    7:"""CROSSTABS
            /TABLES=Campers_vs_staff BY Rate_Israel_programming_at_camp
            /FORMAT=AVALUE TABLES
            /CELLS=COUNT ROW
            /COUNT ROUND CELL.
            """,
    8:"""CROSSTABS
            /TABLES=Campers_vs_staff BY Clusters_Israel_Engagement_factors
            /FORMAT=AVALUE TABLES
            /STATISTICS=CHISQ
            /CELLS=COUNT ROW
            /COUNT ROUND CELL.""",
    9: """ FREQUENCIES VARIABLES=Campers_vs_staff
        /ORDER=ANALYSIS.
        """
}