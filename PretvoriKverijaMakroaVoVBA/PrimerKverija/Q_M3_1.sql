SELECT Q_M3.OBRAZEC AS Expr1, Q_M3.GOD AS Expr2, Q_M3.RBR AS Expr3, [p1]-[p2] AS kontrola1, [p1]-[p3] AS kontrola2, Q_M3.P1 AS Expr4, Q_M3.P2 AS Expr5, Q_M3.P3 AS Expr6
FROM Q_M3
WHERE ((([p1]-[p2])<>0)) Or ((([p1]-[p3])<>0)) Or (((Q_M3.P1) Is Null)) Or (((Q_M3.P2) Is Null)) Or (((Q_M3.P3) Is Null));

