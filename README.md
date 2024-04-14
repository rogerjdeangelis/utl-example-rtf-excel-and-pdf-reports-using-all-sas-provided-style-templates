# uutl-example-rtf-excel-and-pdf-reports-using-all-sas-provided-style-templates
Example rtf excel and pdf reports using all sas provided style templates
    %let pgm=utl-example-rtf-excel-and-pdf-reports-using-all-sas-provided-style-templates;

    Example rtf excel and pdf reports using all sas provided style templates

    github
    https://tinyurl.com/4cy4fyhp
    https://github.com/rogerjdeangelis/uutl-example-rtf-excel-and-pdf-reports-using-all-sas-provided-style-templates

    EXCEL
    https://tinyurl.com/3db5mrf7
    https://github.com/rogerjdeangelis/uutl-example-rtf-excel-and-pdf-reports-using-all-sas-provided-style-templates/blob/main/style.xlsx

    PDF
    https://tinyurl.com/539aru7c
    https://github.com/rogerjdeangelis/uutl-example-rtf-excel-and-pdf-reports-using-all-sas-provided-style-templates/blob/main/style.pdf

    RTF
    https://tinyurl.com/45dcw4u7
    https://github.com/rogerjdeangelis/uutl-example-rtf-excel-and-pdf-reports-using-all-sas-provided-style-templates/blob/main/style.rtf

    Related repo
    https://github.com/rogerjdeangelis/utl_formatting_proc_freq_output_using_a_template
    https://github.com/rogerjdeangelis/utl_sas_classic_graphics_designing_your_greplay_template

    /*----incase you acidentally submit entire program without final changes ----*/
    %stop_submission;

    /*                   _
    (_)_ __  _ __  _   _| |_
    | | `_ \| `_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    */

    /*---- get a list of all sas supplies style templates                    ----*/

    proc template;
       list styles;
    run;

    /*---- create  the meta data for making examples of all style templates  ----*/

    data excel pdf rtf;

       retain sfx ;
       informat style $24.;
       length sty fyl $200;
       input style & @@;

       sfx="excel";
       fyl = compbl("ods excel
                     file='d:/styles/style.xlsx'
                     options(sheet_interval='none');");

       sty = cats("ods excel style=",style,";");
       output excel;

       sfx="pdf";
       fyl = compbl("ods pdf file='d:/styles/style.pdf'
                     startpage=never;");
       sty = cats("ods pdf style=",style,";");
       output pdf;

       sfx="rtf";
       fyl = compbl("ods rtf file='d:/styles/style.rtf'
                     startpage=never ;");
       sty = cats("ods rtf style=",style,";");
       output rtf;

       *if _n_=3 then stop;

    cards4;
    BarrettsBlue        Illuminate         Pearl
    DTree               Journal            PearlJ
    Dove                Journal1a          Plateau
    EGDefault           Journal2           PowerPointDark
    Excel               Journal2a          PowerPointLight
    Excel               Journal3           Printer
    FancyPrinter        Journal3a          Raven
    Festival            Listing            Rtf
    FestivalPrinter     Meadow             Sapphire
    Gantt               MeadowPrinter      SasDocPrinter
    GrayscalePrinter    Minimal            SasWeb
    HTMLBlue            MonochromePrinter  Seaside
    HTMLEncore          Monospace          SeasidePrinter
    Harvest             Moonflower         Snow
    HighContrast        Netdraw            StatDoc
    HighContrastLarge   NoFontDefault      Statistical
    Ignite              Normal             vaDark
    Illuminate          NormalPrinter      vaHighContrast
    Journal             Ocean
    ;;;;
    run;quit;

    /**********************************************************************************************************************************/
    /*                                                                                                                                */
    /* THREE INPUT WORK DATASETS                                                                                                      */
    /*                                                                                                                                */
    /* WORK.EXCEL TOTAL OBS=56                                                                                                        */
    /*                                                                                                                                */
    /*   SFX  STYLE         |                                     FYL                               |                 STY             */
    /*                      |                                                                       |                                 */
    /*  excel BarrettsBlue  | ods excel file='d:/styles/style.xlsx' options(sheet_interval='none'); | ods excel style=BarrettsBlue;   */
    /*  excel Illuminate    | ods excel file='d:/styles/style.xlsx' options(sheet_interval='none'); | ods excel style=Illuminate;     */
    /*  excel Pearl         | ods excel file='d:/styles/style.xlsx' options(sheet_interval='none'); | ods excel style=Pearl;          */
    /*  ...                 |                                                                       |                                 */
    /*                      |                                                                       |                                 */
    /* WORK.PDF total obs=56|                                                                       |                                 */
    /*                      |                                                                       |                                 */
    /*  pdf  BarrettsBlue   | ods pdf file='d:/styles/style.pdf' startpage=never;                   | ods pdf style=BarrettsBlue;     */
    /*  pdf  Illuminate     | ods pdf file='d:/styles/style.pdf' startpage=never;                   | ods pdf style=Illuminate;       */
    /*  pdf  Pearl          | ods pdf file='d:/styles/style.pdf' startpage=never;                   | ods pdf style=Pearl;            */
    /* ...                  |                                                                       |                                 */
    /*                      |                                                                       |                                 */
    /* WORK.RTF total obs=56|                                                                       |                                 */
    /*                      |                                                                       |                                 */
    /*  rtf  BarrettsBlue   | ods rtf file='d:/styles/style.rtf' startpage=never ;                  | ods rtf style=BarrettsBlue;     */
    /*  rtf  Illuminate     | ods rtf file='d:/styles/style.rtf' startpage=never ;                  | ods rtf style=Illuminate;       */
    /*  rtf  Pearl          | ods rtf file='d:/styles/style.rtf' startpage=never ;                  | ods rtf style=Pearl;            */
    /* ...                  |                                                                       |                                 */
    /**********************************************************************************************************************************/

    /*
     _ __  _ __ ___   ___ ___  ___ ___
    | `_ \| `__/ _ \ / __/ _ \/ __/ __|
    | |_) | | | (_) | (_|  __/\__ \__ \
    | .__/|_|  \___/ \___\___||___/___/
    |_|
    */

    /*----                                                                   ----*/
    /*---- dosubl faster then call execute or as fast (who cares anyway)     ----*/
    /*---- dosubl cleaner and more easily maintained than call execute       ----*/
    /*---- Eliminates all the quoting needed needed with call execute        ----*/
    /*----                                                                   ----*/

    ods _all_ close;

    /*----  close any hidden open files                                      ----*/
    %utl_close;

    %utlfkil(d:/styles/style.pdf);
    %utlfkil(d:/styles/style.xlsx);
    %utlfkil(d:/styles/style.rtf);

    title;
    footnote;

    data _null_;

      set
          excel
          pdf
          rtf
      ;

      by fyl;

      call symputx('style',style);
      call symputx('fyl',fyl);
      call symputx('sty',sty);
      call symputx('sfx',sfx);

      if first.fyl then do;
          /*---- ods pdf file='d:/styles/style.pdf' startpage=never          ----*/
          rc=dosubl('&fyl');
      end;

      rc=dosubl('
          /*---- ods pdf style=BarrettsBlue;                                 ----*/
         &sty;
         proc report data=sashelp.class(obs=3)
               style(report)={fontsize=18pt}
               style(header)={font_size=15pt}
               style(column)={font_size=14pt}
           ;
           cols ("&style" name sex age);
         run;quit;
         ');

      if last.fyl then do;
        rc=dosubl('
          ods &sfx close;
          ');
      end;

    run;quit;

    ods listing;
    run;quit;

    /*           _               _
      ___  _   _| |_ _ __  _   _| |_
     / _ \| | | | __| `_ \| | | | __|
    | (_) | |_| | |_| |_) | |_| | |_
     \___/ \__,_|\__| .__/ \__,_|\__|
                    |_|
    */


    EXCEL
    https://tinyurl.com/3db5mrf7
    https://github.com/rogerjdeangelis/uutl-example-rtf-excel-and-pdf-reports-using-all-sas-provided-style-templates/blob/main/style.xlsx

    PDF
    https://tinyurl.com/539aru7c
    https://github.com/rogerjdeangelis/uutl-example-rtf-excel-and-pdf-reports-using-all-sas-provided-style-templates/blob/main/style.pdf

    RTF
    https://tinyurl.com/45dcw4u7
    https://github.com/rogerjdeangelis/uutl-example-rtf-excel-and-pdf-reports-using-all-sas-provided-style-templates/blob/main/style.rtf

    /*              _
      ___ _ __   __| |
     / _ \ `_ \ / _` |
    |  __/ | | | (_| |
     \___|_| |_|\__,_|

    */
