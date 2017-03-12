/*
DROP TABLE FOOTBALL_STADIUMS;
DROP TABLE FOOTBALL_TEAMS;
DROP TABLE FOOTBALL_PLAYERS;
DROP TABLE FOOTBALL_MATCHES;*/

/*alter table FOOTBALL_TEAMS add(points NUMBER(8));
commit;*//*
alter table FOOTBALL_PLAYERS modify(number_player NUMBER(8));
commit;*//*
alter table football_matches add(ball_first NUMBER(8), ball_second NUMBER(8));
commit;*//*
alter table football_matches rename column id_team_sk to id_team_st;
commit;*/


/*
CREATE TABLE FOOTBALL_STADIUMS( 
id_stadium NUMBER(8) NOT NULL PRIMARY KEY,
title VARCHAR2(40) NOT NULL,
city VARCHAR2(40) NOT NULL,
capacity_stadium NUMBER(8) NOT NULL
);

CREATE TABLE FOOTBALL_TEAMS( 
id_team NUMBER(8) NOT NULL PRIMARY KEY,
name_team varchar2(80) NOT NULL, 
base varchar2(40) NOT NULL,
coach varchar2(80) NOT NULL,
position_team NUMBER(8) NOT NULL,
win NUMBER(8),
defeat NUMBER(8),
draw NUMBER(8)
);

CREATE TABLE FOOTBALL_PLAYERS( 
id_player NUMBER(8) NOT NULL PRIMARY KEY,
id_team_fk NUMBER(8) NOT NULL REFERENCES FOOTBALL_TEAMS(ID_TEAM), 
fio varchar2(80) NOT NULL,
age NUMBER(8) NOT NULL,
number_player DATE NOT NULL,
role_player VARCHAR2(40) NOT NULL,
CHECK(role_player IN ('вратарь','защитник','нападающий')) 
);

CREATE TABLE FOOTBALL_MATCHES( 
ID_MATCH NUMBER(8) NOT NULL PRIMARY KEY,
id_team_ft NUMBER(8) NOT NULL REFERENCES FOOTBALL_TEAMS(ID_TEAM), 
id_team_sk NUMBER(8) NOT NULL REFERENCES FOOTBALL_TEAMS(ID_TEAM),
id_stadium NUMBER(8) NOT NULL REFERENCES FOOTBALL_STADIUMS(ID_STADIUM),
cost_ticket NUMBER(8) NOT NULL,
date_match DATE NOT NULL
);
*/
