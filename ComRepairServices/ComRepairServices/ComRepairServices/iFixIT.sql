Create Database iFixIT;
use iFixIT;
-- Create Users Table
CREATE TABLE users (
    user_id INT PRIMARY KEY AUTO_INCREMENT,
    username VARCHAR(255) NOT NULL,
    email VARCHAR(255) NOT NULL
    -- Add more columns as needed
);

-- Create Product List Table
CREATE TABLE product_list (
    product_id INT PRIMARY KEY AUTO_INCREMENT,
    product_name VARCHAR(255) NOT NULL,
    price DECIMAL(10, 2) NOT NULL
    -- Add more columns as needed
);

-- Create Service List Table
CREATE TABLE service_list (
    service_id INT PRIMARY KEY AUTO_INCREMENT,
    service_name VARCHAR(255) NOT NULL,
    price DECIMAL(10, 2) NOT NULL
    -- Add more columns as needed
)

-- Create Transaction List Table
CREATE TABLE transaction_list (
    transaction_id INT PRIMARY KEY AUTO_INCREMENT,
    user_id INT,
    transaction_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    -- Add more columns as needed
)

-- Create Transaction Products Table
CREATE TABLE transaction_products (
    transaction_id INT,
    product_id INT,
    quantity INT NOT NULL,
    PRIMARY KEY (transaction_id, product_id),
    FOREIGN KEY (transaction_id) REFERENCES transaction_list(transaction_id),
    FOREIGN KEY (product_id) REFERENCES product_list(product_id)
)

-- Create Transaction Services Table
CREATE TABLE transaction_services (
    transaction_id INT,
    service_id INT,
    PRIMARY KEY (transaction_id, service_id),
    FOREIGN KEY (transaction_id) REFERENCES transaction_list(transaction_id),
    FOREIGN KEY (service_id) REFERENCES service_list(service_id)
);
