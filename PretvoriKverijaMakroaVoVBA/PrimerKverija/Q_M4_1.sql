SELECT Q_M4.GOD AS Expr1, Q_M4.RBR AS Expr2, [p1]-[p2] AS kontrola1, [p1]-[p3] AS kontrola2, Q_M4.P1 AS Expr3, Q_M4.P2 AS Expr4, Q_M4.P3 AS Expr5
FROM Q_M4
WHERE ((([p1]-[p2])<>0)) Or ((([p1]-[p3])<>0)) Or (((Q_M4.P1) Is Null)) Or (((Q_M4.P2) Is Null)) Or (((Q_M4.P3) Is Null));

