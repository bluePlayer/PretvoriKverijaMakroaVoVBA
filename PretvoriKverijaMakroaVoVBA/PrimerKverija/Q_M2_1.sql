SELECT Q_M2.OBRAZEC AS Expr1, Q_M2.GOD AS Expr2, Q_M2.RBR AS Expr3, [p1]-[p2] AS kontrola1, [p1]-[p3] AS kontrola2, [p1]-[p4] AS kontrola3, Q_M2.P1 AS Expr4, Q_M2.P2 AS Expr5, Q_M2.P3 AS Expr6, Q_M2.P4 AS Expr7
FROM Q_M2
WHERE ((([p1]-[p2])<>0)) Or ((([p1]-[p3])<>0)) Or ((([p1]-[p4])<>0)) Or (((Q_M2.P1) Is Null)) Or (((Q_M2.P2) Is Null)) Or (((Q_M2.P3) Is Null)) Or (((Q_M2.P4) Is Null));

