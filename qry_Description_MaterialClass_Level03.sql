SELECT lst_FISHThesaurus.TERM, lst_FISHThesaurus.TERM_DEFF
FROM lst_FISHThesaurus
WHERE (((lst_FISHThesaurus.TERM_ISTOP)=False) AND ((lst_FISHThesaurus.CLA_GR_UID)=73) AND ((lst_FISHThesaurus.LINK_TERM)=[Forms]![nav_LandingPage]![NavigationSubform]![Description_MaterialClass_Level02]))
ORDER BY lst_FISHThesaurus.TERM;
