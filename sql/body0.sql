SELECT t0.h0_txt,
       t0.h1_txt,
       t0.h2_txt,
       t0.t_value
  FROM t_0 t0
 WHERE t0.ym = (SELECT MAX(t1.ym)) FROM t_0 t1 WHERE t1.ym <= {});
