CREATE TABLE IF NOT EXISTS tb_clientes  (
   id  INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
   cpf_cnpj  varchar(20) DEFAULT NULL,
   nome  varchar(50) NOT NULL,
   usuario  varchar(15) DEFAULT NULL,
   celular  varchar(15) NOT NULL,
   email  varchar(60) NOT NULL,
   endereco  varchar(150) DEFAULT NULL,
   numero  varchar(10) DEFAULT NULL,
   complemento  varchar(10) DEFAULT NULL,
   bairro  varchar(100) DEFAULT NULL,
   cidade  varchar(100) DEFAULT NULL,
   estado  varchar(2) DEFAULT NULL,
   cep  varchar(9) DEFAULT NULL,
   ultcompra  date DEFAULT NULL,
   obs  varchar(250) DEFAULT NULL,
   operador  int DEFAULT NULL,
   datatual  date DEFAULT NULL
 ) ;

CREATE TABLE  tb_filamentos  (
   id  INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
   marca  varchar(100)   DEFAULT NULL,
   tipo  varchar(100)  DEFAULT NULL,
   cor  varchar(100)  DEFAULT NULL,
   valor  decimal(10,0) NOT NULL,
   qtde_estoque  int NOT NULL,
   operador  int DEFAULT NULL,
   datatual  date DEFAULT NULL
  );

CREATE TABLE  tb_fornecedores  (
   id  INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
   cpf_cnpj  varchar(20) DEFAULT NULL,
   nome  varchar(45) DEFAULT NULL,
   endereco  varchar(145) DEFAULT NULL,
   email  varchar(145) DEFAULT NULL,
   celular  varchar(20) DEFAULT NULL,
   operador  int DEFAULT NULL,
   datatual  datetime DEFAULT NULL  
) ;

CREATE TABLE  tb_impressoras  (
   id  INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
   marca  varchar(250)  DEFAULT NULL,
   modelo  varchar(250) DEFAULT NULL,
   qtrolos  int NOT NULL,
   ocupada  varchar(1)  DEFAULT NULL,
   operador  int DEFAULT NULL,
   datatual  date DEFAULT NULL
  
) ;

CREATE TABLE  tb_lojas  (
   id  INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
   nome  varchar(50) NOT NULL,
   endereco  varchar(150) DEFAULT NULL,
   CPF_CNPJ  varchar(15) DEFAULT NULL,
   telefone  varchar(15) DEFAULT NULL,
   senha  varchar(1) DEFAULT NULL,
   operador  int DEFAULT NULL,
   datatual  date DEFAULT NULL
  
); 

CREATE TABLE  tb_niveis  (
   id  INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
   descricao  varchar(20) NOT NULL,
   operador  int DEFAULT NULL,
   datatual  date DEFAULT NULL
 
);

CREATE TABLE  tb_filamentos  (
   id  integer NOT NULL PRIMARY KEY AUTOINCREMENT,
   marca  varchar(100)  DEFAULT NULL,
   tipo  varchar(100)  DEFAULT NULL,
   cor  varchar(100)  DEFAULT NULL,
   valor  decimal(10,0) NOT NULL,
   qtde_estoque  int NOT NULL,
   operador  int DEFAULT NULL,
   datatual  date DEFAULT NULL
 );

CREATE TABLE  tb_pedidos  (
   id_venda  INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
   id_cliente  int NOT NULL,
   descricao  varchar(250) NOT NULL,
   preco  float NOT NULL,
   quantidade  int NOT NULL,
   total_venda  float NOT NULL,
   situacao  int NOT NULL,
   dataCompra  date NOT NULL,
   dataInicioProd  date DEFAULT NULL,
   dataPrevisao  date NOT NULL,
   dataFinaliza  date DEFAULT NULL,
   dataEntrega  date DEFAULT NULL,
   operador  int DEFAULT NULL,
   datatual  date DEFAULT NULL
 );

CREATE TABLE  tb_pedxfilamento  (
   id_pedido  int NOT NULL,
   id_filamento  int NOT NULL,
   operador  int DEFAULT NULL,
   datatual  date DEFAULT NULL
 ) ;
 
CREATE INDEX idx_pedfilamento 
ON tb_pedxfilamento (id_pedido, id_filamento);

CREATE TABLE  tb_situacao  (
   id_situacao  INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
   descricao  varchar(10)  NOT NULL,
   operador  int DEFAULT NULL,
   datatual  date DEFAULT NULL
  
) ;

CREATE TABLE  tb_usuarios  (
   id  INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
   nome  varchar(50) NOT NULL,
   login  varchar(10) NOT NULL,
   email  varchar(100) DEFAULT NULL,
   celular  varchar(15) DEFAULT NULL,
   senha  varchar(8) NOT NULL,
   nivel  tinyint NOT NULL,
   salario  decimal(12,2) DEFAULT NULL,
   comissao  decimal(5,2) DEFAULT NULL,
   ativo  tinyint DEFAULT NULL,
   operador  int DEFAULT NULL,
   datatual  date DEFAULT NULL
  ) ;


CREATE TABLE  tb_pedximpres  (
   id_pedido  int NOT NULL,
   id_impressora  int NOT NULL,
   operador  int DEFAULT NULL,
   datatual  date DEFAULT NULL,
  KEY  id_pedimpress  ( id_pedido , id_impressora )
) ;

CREATE INDEX idx_pedimpres 
ON tb_pedximpres (id_pedido, id_impressora);

