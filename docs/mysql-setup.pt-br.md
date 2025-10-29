# Guia: MySQL central para múltiplos PCs

Este passo a passo explica como disponibilizar um único servidor MySQL para que qualquer PC que rode o MedAlgoApp se conecte diretamente. Inclui opções on‑premises (Windows), Docker e cloud, além de segurança e testes.

> Recomendação: Evite expor MySQL diretamente na Internet. Prefira rede local/VPN ou um serviço gerenciado (Azure/AWS/GCP) com firewall/SSL. Se a Internet for necessária, exija SSL, IP allowlist e senha forte.

## 1) Escolha de arquitetura

- Servidor central: 1 máquina (Windows ou Linux) rodando MySQL, acessível pela LAN/VPN.
- Cloud gerenciado: Azure Database for MySQL / AWS RDS / Cloud SQL — simplifica backup, SSL e disponibilidade.
- Docker: Executa MySQL em um container fixando porta 3306 e volume para dados.

Todos os clientes apontam para o mesmo host:porta e usam o mesmo schema.

## 2) Instalação no Windows (on‑premise)

1. Instale o MySQL Community Server.
2. Edite o arquivo de configuração (my.ini):
   - Habilite escuta externa:
     - `bind-address = 0.0.0.0`
   - Garanta que rede está habilitada (não use `skip_networking = 1`).
3. Abra a porta no Firewall (como admin):

```cmd
netsh advfirewall firewall add rule name="MySQL 3306" dir=in action=allow protocol=TCP localport=3306
```

4. Reinicie o serviço MySQL:

```cmd
net stop MySQL80
net start MySQL80
```

5. Crie base e usuário com acesso remoto (ajuste host/IP conforme seu cenário):

```sql
CREATE DATABASE medalgo CHARACTER SET utf8mb4 COLLATE utf8mb4_0900_ai_ci;
CREATE USER 'medalgo_app'@'10.0.%' IDENTIFIED BY 'TroqueEstaSenha!';
GRANT ALL PRIVILEGES ON medalgo.* TO 'medalgo_app'@'10.0.%';
FLUSH PRIVILEGES;
```

> Dica: Em redes domésticas/Internet, use `user'@'seu_ip_publico'` ou um range específico. Evite `'%'` (qualquer IP) sempre que possível.

## 3) Instalação via Docker

```cmd
docker run -d --name mysql-medalgo -p 3306:3306 ^
  -e MYSQL_ROOT_PASSWORD=TroqueRoot! ^
  -e MYSQL_DATABASE=medalgo ^
  -e MYSQL_USER=medalgo_app ^
  -e MYSQL_PASSWORD=TroqueEstaSenha! ^
  -v %cd%\mysql_data:/var/lib/mysql ^
  mysql:8.4
```

- Certifique-se de abrir a porta 3306 conforme a seção de firewall.

## 4) Criar o schema (tabelas)

O script abaixo replica a estrutura funcional hoje usada no SQLite (campos de data como texto para compatibilidade). Ajuste tipos para `DATETIME` se quiser normalizar.

Arquivo sugerido: `docs/mysql-schema.sql`

```sql
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
```

## 5) Teste de conectividade a partir de um cliente

1. Instale o cliente MySQL ou use Docker para testar:

```cmd
mysql -h SEU_HOST -P 3306 -u medalgo_app -p
```

2. Verifique o schema:

```sql
SHOW DATABASES; USE medalgo; SHOW TABLES; SELECT 1;
```

Se não conectar: confira firewall/porta, IP, DNS, regras de `GRANT` e NAT.

## 6) Conexão pelo aplicativo (.NET)

O MedAlgoApp hoje usa SQLite local. Para usar MySQL central, será preciso trocar o provedor no código e ajustar alguns comandos SQL.

- Pacote NuGet recomendado: `MySqlConnector`
- Connection string (exemplo):

```
Server=SEU_HOST;Port=3306;Database=medalgo;User ID=medalgo_app;Password=TroqueEstaSenha!;SslMode=Preferred;Pooling=true;MinimumPoolSize=0;MaximumPoolSize=50;
```

- Segurança: em Internet pública, prefira `SslMode=Required` e valide certificado.

### Diferenças de SQL a considerar

- `AUTOINCREMENT` (SQLite) → `AUTO_INCREMENT` (MySQL)
- `INSERT ... ON CONFLICT DO UPDATE` (SQLite) → `INSERT ... ON DUPLICATE KEY UPDATE` (MySQL)
- Tipos de data: hoje usamos `TEXT` com formato `dd/MM/yyyy HH:mm:ss`; você pode migrar para `DATETIME`.

### Estratégia sugerida no app

- Adicionar uma chave de configuração `DbProvider` ("sqlite" | "mysql") e `MySqlConnectionString`.
- No código, selecionar o provedor e usar SQL específico por provedor.
- Alternativa: adotar EF Core para abstrair diferenças (mais trabalho inicial, menos SQL manual).

Se quiser, posso implementar essa troca de provedor (SQLite/MySQL) no projeto e uma tela de configuração para salvar a connection string.

## 7) Boas práticas de segurança

- Nunca use `root` no app. Crie usuário de aplicação com permissões mínimas.
- Restrinja o host (`'usuario'@'ip_ou_range'`).
- Use senhas fortes e rotacione periodicamente.
- Habilite SSL/TLS e, se possível, VPN ou IP allowlist.
- Faça backup regular e monitore logs de acesso.

## 8) Alternativa recomendada (API em vez de DB direto)

Para cenários via Internet, considere criar uma pequena API (REST) entre o app e o banco. Vantagens:
- Evita expor o MySQL ao mundo.
- Aplica regras de negócio/validações no servidor.
- Possibilita cache, auditoria, e escalabilidade.

Posso gerar um template de API .NET minimal que espelha as operações já usadas pelo app.
