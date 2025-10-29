-- Schema MySQL para MedAlgoApp
-- Execute isto no banco 'medalgo'

CREATE TABLE IF NOT EXISTS config (
  id INT AUTO_INCREMENT PRIMARY KEY,
  chave VARCHAR(255) UNIQUE,
  valor TEXT
);

CREATE TABLE IF NOT EXISTS chaves (
  id INT AUTO_INCREMENT PRIMARY KEY,
  nome VARCHAR(255) UNIQUE,
  num_copias INT,
  descricao TEXT
);

CREATE TABLE IF NOT EXISTS reservas (
  id INT AUTO_INCREMENT PRIMARY KEY,
  chave VARCHAR(255),
  aluno VARCHAR(255),
  professor VARCHAR(255),
  data_hora VARCHAR(19),
  em_uso TINYINT(1),
  termo TEXT,
  devolvido TINYINT(1),
  data_devolucao VARCHAR(19)
);

CREATE TABLE IF NOT EXISTS relatorio (
  id INT AUTO_INCREMENT PRIMARY KEY,
  chave VARCHAR(255),
  aluno VARCHAR(255),
  professor VARCHAR(255),
  data_hora VARCHAR(19),
  data_devolucao VARCHAR(19),
  tempo_com_chave VARCHAR(255),
  termo TEXT
);
