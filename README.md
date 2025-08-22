# Thread-Safe Spreadsheet Simulator

## Overview
Backend-focused C# project that implements a shared in-memory spreadsheet with **fine-grained locking** and **concurrent workloads**.  
The GUI (WinForms) is minimal and serves only as a demo surface.  
Main focus: **OOP design, synchronization, and multithreading**.

## Features
- **Thread-Safe Backend**: Fine-grained locking mechanism ensures safe concurrent updates.  
- **Multithreading**: Multiple simulated users interact with the spreadsheet simultaneously.  
- **OOP Design**: Encapsulated spreadsheet class with clear API (`setCell`, `getCell`, `searchInRow`, etc).  
- **Simulator**: Runs randomized concurrent workloads to stress-test locking.  
- **CSV Support**: Load and save spreadsheets to `.csv` files.  

## Technologies
- C# (.NET 6/7)
- Windows Forms (basic UI for demo)
- Multithreading & Synchronization (Locks, Mutex)
- OOP

## Why This Project
Designed to demonstrate backend engineering principles:  
- Safe handling of **shared mutable state**.  
- Correctness under concurrent operations.  
- Efficient synchronization without deadlocks.  

## How to Run
1. Clone the repository:
   ```bash
   git clone https://github.com/<your-username>/ThreadSafe-Spreadsheet.git
   ```
2. Open `SpreadsheetApp.csproj` in Visual Studio.  
3. Build & Run.

## License
MIT
