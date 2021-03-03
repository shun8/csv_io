SELECT t0.h_txt,
       t0.h_order
  FROM h_2 t0
 WHERE t0.ym = (SELECT MAX(t1.ym) FROM h_2 t1 WHERE t1.ym <= %(yyyymm)s);
