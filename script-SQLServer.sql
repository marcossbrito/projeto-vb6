CREATE DATABASE db_sistemacorretor
ON PRIMARY (
NAME=db_sistemacorretor,
FILENAME='C:\SQL\db_sistemacorretor.MDF',
SIZE=10MB,
MAXSIZE=100MB,
filegrowth=10%
);


use db_sistemacorretor;


CREATE TABLE corretor(
id INT PRIMARY KEY IDENTITY(1,1),
codigo VARCHAR(7) NOT NULL,
nome VARCHAR(50) NOT NULL,
cpf VARCHAR(11) NOT NULL
);


CREATE TABLE estado(
id INT PRIMARY KEY IDENTITY(1,1),
uf VARCHAR(2) NOT NULL
);


CREATE TABLE cidade(
id INT PRIMARY KEY IDENTITY(1,1),
nome VARCHAR(50) NOT NULL,
id_uf INT NOT NULL,

CONSTRAINT fk_id_uf FOREIGN KEY (id_uf)
REFERENCES estado (id)
);


CREATE TABLE cliente(
id INT PRIMARY KEY IDENTITY,
nome VARCHAR(50) NOT NULL,
cpf VARCHAR(11) NOT NULL,
endereco VARCHAR(90)NOT NULL,
ativo BIT NOT NULL,
id_uf INT NOT NULL,
id_cidade INT NOT NULL,
id_corretor INT NOT NULL,

CONSTRAINT fk_id_uf_cliente FOREIGN KEY (id_uf)
REFERENCES estado (id),

CONSTRAINT fk_id_cidade_cliente FOREIGN KEY (id_cidade)
REFERENCES cidade (id),

CONSTRAINT fk_id_corretor_cliente FOREIGN KEY (id_corretor)
REFERENCES corretor (id)
);


INSERT INTO estado (uf) VALUES ('SP');
INSERT INTO estado (uf) VALUES ('MG');
INSERT INTO estado (uf) VALUES ('BA');
INSERT INTO estado (uf) VALUES ('RJ');


INSERT INTO cidade (nome, id_uf) VALUES ('Ribeirão Preto', 1);
INSERT INTO cidade (nome, id_uf) VALUES ('Campinas', 1);
INSERT INTO cidade (nome, id_uf) VALUES ('Uberaba', 2);
INSERT INTO cidade (nome, id_uf) VALUES ('Salvador', 3);
INSERT INTO cidade (nome, id_uf) VALUES ('Rio de Janeiro', 4);