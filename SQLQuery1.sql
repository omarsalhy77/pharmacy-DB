 use  pharmacyDB

/* =========================
   Customers
========================= */
CREATE TABLE Customers (
    CustomerID INT PRIMARY KEY IDENTITY(1,1),
    FirstName VARCHAR(50) NOT NULL,
    LastName VARCHAR(50) NOT NULL,
    Phone VARCHAR(15) NOT NULL UNIQUE
);

/* =========================
   Staff
========================= */
CREATE TABLE Staff (
    StaffID INT PRIMARY KEY IDENTITY(1,1),
    FirstName VARCHAR(50) NOT NULL,
    LastName VARCHAR(50) NOT NULL,
    Email VARCHAR(100) NOT NULL UNIQUE,
    Phone VARCHAR(15) NOT NULL UNIQUE,
    Role VARCHAR(30) NOT NULL,
    Address VARCHAR(150),
    Salary DECIMAL(10,2) NOT NULL CHECK (Salary > 0),
    
    CONSTRAINT CHK_StaffRole 
    CHECK (Role IN ('Pharmacist', 'Cashier', 'Manager'))
);

/* =========================
   Suppliers
========================= */
CREATE TABLE Suppliers (
    SupplierID INT PRIMARY KEY IDENTITY(1,1),
    SupplierName VARCHAR(100) NOT NULL,
    Phone VARCHAR(15) NOT NULL UNIQUE,
    Address VARCHAR(150)
);

/* =========================
   Categories
========================= */
CREATE TABLE Categories (
    CategoryID INT PRIMARY KEY IDENTITY(1,1),
    CategoryName VARCHAR(50) NOT NULL UNIQUE
);

/* =========================
   Medicines
========================= */
CREATE TABLE Medicines (
    MedicineID INT PRIMARY KEY IDENTITY(1,1),
    MedicineName VARCHAR(100) NOT NULL,
    SupplierID INT NOT NULL,
    CategoryID INT NOT NULL,
    Price DECIMAL(10,2) NOT NULL CHECK (Price > 0),
    Quantity INT NOT NULL CHECK (Quantity >= 0),
    ExpiryDate DATE NOT NULL,

    CONSTRAINT FK_Medicine_Supplier
    FOREIGN KEY (SupplierID) REFERENCES Suppliers(SupplierID),

    CONSTRAINT FK_Medicine_Category
    FOREIGN KEY (CategoryID) REFERENCES Categories(CategoryID)
);

/* =========================
   PurchaseOrders
========================= */
CREATE TABLE PurchaseOrders (
    OrderID INT PRIMARY KEY IDENTITY(1,1),
    SupplierID INT NOT NULL,
    OrderDate DATE DEFAULT GETDATE(),

    CONSTRAINT FK_Order_Supplier
    FOREIGN KEY (SupplierID) REFERENCES Suppliers(SupplierID)
);

/* =========================
   PurchaseDetails
========================= */
CREATE TABLE PurchaseDetails (
    OrderID INT NOT NULL,
    MedicineID INT NOT NULL,
    Quantity INT NOT NULL CHECK (Quantity > 0),
    UnitPrice DECIMAL(10,2) NOT NULL CHECK (UnitPrice > 0),

    PRIMARY KEY (OrderID, MedicineID),

    CONSTRAINT FK_PurchaseDetails_Order
    FOREIGN KEY (OrderID) REFERENCES PurchaseOrders(OrderID),

    CONSTRAINT FK_PurchaseDetails_Medicine
    FOREIGN KEY (MedicineID) REFERENCES Medicines(MedicineID)
);

/* =========================
   Sales
========================= */
CREATE TABLE Sales (
    SaleID INT PRIMARY KEY IDENTITY(1,1),

    CustomerID INT NOT NULL,
    StaffID INT NOT NULL,

    SaleDate DATETIME NOT NULL
    CONSTRAINT DF_Sales_SaleDate DEFAULT GETDATE(),

    TotalAmount DECIMAL(10,2) NOT NULL
    CHECK (TotalAmount >= 0),

    Discount DECIMAL(10,2) NOT NULL
    DEFAULT 0
    CHECK (Discount >= 0),

    CONSTRAINT FK_Sales_Customer
    FOREIGN KEY (CustomerID)
    REFERENCES Customers(CustomerID),

    CONSTRAINT FK_Sales_Staff
    FOREIGN KEY (StaffID)
    REFERENCES Staff(StaffID)
);
GO
CREATE TABLE SaleDetails (
    SaleID INT NOT NULL,
    MedicineID INT NOT NULL,

    Quantity INT NOT NULL
    CHECK (Quantity > 0),

    UnitPrice DECIMAL(10,2) NOT NULL
    CHECK (UnitPrice > 0),

    PRIMARY KEY (SaleID, MedicineID),

    CONSTRAINT FK_SaleDetails_Sale
    FOREIGN KEY (SaleID)
    REFERENCES Sales(SaleID),

    CONSTRAINT FK_SaleDetails_Medicine
    FOREIGN KEY (MedicineID)
    REFERENCES Medicines(MedicineID)
);
GO