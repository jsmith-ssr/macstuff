%macro acrf_xfdf_xlsx(basepath=);

%local t_start t_end;
%let t_start=%sysfunc(datetime());

%put NOTE: ============================================;
%put NOTE: ACRF XFDF TO XLSX;
%put NOTE: %sysfunc(datetime(),datetime19.);
%put NOTE: ============================================;

%let xfdfmap = &basepath.\xfdf_map.map;
%let crf_xfdf = &basepath.\aCRF.xfdf;
%let outxlsx = &basepath.\acrf.xlsx;

/* XML load */
filename xfdfmap "&xfdfmap";
filename crf_xfdf "&crf_xfdf";

libname crf_xfdf xml xmlmap=xfdfmap access=readonly;

proc copy in=crf_xfdf out=work memtype=data;
run;

%put NOTE: XML import complete.;

/* ---------- High-performance streaming transform ---------- */

data _final(keep=page text domain color);

    length domain mydomain $200;
    retain domain rows kept;

    if _n_=1 then do;
        rows=0;
        kept=0;

        declare hash dedup(hashexp:20);
        dedup.definekey('page','color','text');
        dedup.definedone();

        put "NOTE: Hash dedup initialized.";
    end;

    set annots end=done;

    rows+1;
    if mod(rows,50000)=0 then
        put "NOTE: Processed rows=" rows;

    page+1;

    if text='[NOT SUBMITTED]' then return;

    /* dedup */
    if dedup.check()=0 then return;
    dedup.add();

    /* --- ORIGINAL DOMAIN LOGIC RESTORED --- */
    if prxmatch('/[()]/', text) then do;
        mydomain = strip(prxchange('s/\([^)]*\)|[()]//', -1, text));
        if not missing(mydomain) then domain=mydomain;
        return;  /* do NOT output domain marker rows */
    end;

    color=cats('cx',substr(color,2));

    kept+1;
    output;

    if done then do;
        put "NOTE: Total rows read =" rows;
        put "NOTE: Rows exported   =" kept;
    end;

run;

/* Final presentation sort only */
proc sort data=_final;
    by page color domain text;
run;

%put NOTE: Transform complete.;

/* ---------- Export ---------- */

ods excel file="&outxlsx";

proc report data=_final nowd;
    column page text domain color;

    define page   / display "PAGENO";
    define text   / display "TEXT";
    define domain / display "DOMAIN";
    define color  / noprint;

    compute color;
        call define(_row_, 'STYLE', 'STYLE={background='||color||'}');
    endcomp;
run;

ods excel close;

%let t_end=%sysfunc(datetime());

%put NOTE: ============================================;
%put NOTE: EXPORT COMPLETE;
%put NOTE: Duration = %sysevalf(&t_end-&t_start) seconds;
%put NOTE: %sysfunc(datetime(),datetime19.);
%put NOTE: ============================================;

%mend acrf_xfdf_xlsx;

/*Example to call*/
/*%acrf_xfdf_xlsx(basepath=to_xfdf_patg.xfdf);*/
