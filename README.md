# CPReservation

Online reservation system for **Chilli Padi Restaurant**, built using **Google Apps Script**.

This system allows customers to:
- Make table reservations online
- Receive email confirmation
- Prevent overbooking based on pax limit
- Automatically block closed dates and time ranges

## Features

- Date & time-based reservation system
- Pax control (Adults + Children limit)
- Double-booking prevention using LockService
- Automatic email confirmation
- Date and time blocking via Google Sheet
- Reservation records stored in Google Sheets

## Reservation Flow

1. Customer submits reservation form
2. System checks:
   - If date is fully blocked
   - If time range is blocked
   - If pax exceeds maximum capacity
3. LockService prevents race condition
4. Reservation is saved to Google Sheets
5. Confirmation email is sent to customer
