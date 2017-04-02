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
        """
    # ,
    # 2:"""CTABLES
    #     A_D_learn_jewish_text_independently Grade_Level ORDER=A KEY=VALUE EMPTY=INCLUDE MISSING=EXCLUDE
    #                     /CRITERIA CILEVEL=95
    #                     /TITLES
    #                     TITLE=' Exhibit 22: Students Assessment of their Hebrew for Text Study and Prayer Abilities.'.""")
}