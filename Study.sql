CREATE TABLE sqlite_stat1(tbl,idx,stat);
CREATE TABLE study (
    "id" INTEGER PRIMARY KEY,
    "fio" VARCHAR(80),
    "gruppa" VARCHAR(80)
);
CREATE TABLE delays (
    "id" INTEGER PRIMARY KEY,
    "fio" VARCHAR(80),
    "gruppa" VARCHAR(40),
    "dod" REAL,
    "hours_u" INTEGER,
    "hours_n" INTEGER
);
CREATE TABLE groups (
    "id" INTEGER PRIMARY KEY,
    "gruppa" VARCHAR(20) NOT NULL
);
