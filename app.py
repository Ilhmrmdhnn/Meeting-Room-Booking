import streamlit as st
import pandas as pd
from datetime import datetime

# Initialize an empty DataFrame if no file exists
def load_bookings():
    try:
        bookings = pd.read_excel('bookings.xlsx', engine='openpyxl')
        # Ensure 'Tanggal' column is a datetime object
        bookings['Tanggal'] = pd.to_datetime(bookings['Tanggal']).dt.date
        return bookings
    except FileNotFoundError:
        return pd.DataFrame(columns=['Nama', 'SBU', 'Ruangan', 'Tanggal', 'Time In', 'Time Out', 'Status'])

# Function to save bookings to Excel
def save_bookings(df):
    df.to_excel('bookings.xlsx', index=False, engine='openpyxl')

# Load or create bookings data
bookings_df = load_bookings()

# Streamlit inputs
st.title('Meeting Room Booking OMO SML')

nama = st.text_input('Nama')
sbu = st.text_input('SBU')
ruangan = st.selectbox('Ruangan', ['Loyal', 'Commitment', 'Integrity'])
tanggal = st.date_input('Tanggal')
time_in = st.selectbox('Time In', [f'{h:02d}:{m:02d}' for h in range(6, 19) for m in (0, 30)])
time_out = st.selectbox('Time Out', [f'{h:02d}:{m:02d}' for h in range(6, 19) for m in (0, 30)])

# Convert the 'Tanggal' input to a date object
tanggal = tanggal

# Check if the room is available for the selected time slot
booking_exists = (
    bookings_df[
        (bookings_df['Ruangan'] == ruangan) & 
        (bookings_df['Tanggal'] == tanggal) & 
        (bookings_df['Time In'] <= time_out) & 
        (bookings_df['Time Out'] >= time_in)
    ].shape[0] > 0
)

# Define status and color
status = 'booked' if booking_exists else 'available'
color = 'red' if status == 'booked' else 'green'

# Display status with color coding
st.markdown(f'<p style="color: {color}; font-size: 20px; font-weight: bold;">Status: {status.capitalize()}</p>', unsafe_allow_html=True)

# Save booking or show error if not available
if st.button('Book Room', key='book_room'):
    # Check if all required fields are filled
    if not nama or not sbu or not ruangan or not tanggal or not time_in or not time_out:
        st.warning('Please fill in all the required fields!')
    else:
        if booking_exists:
            st.error('The selected room and time slot is already booked.')
        else:
            new_booking = pd.DataFrame({
                'Nama': [nama],
                'SBU': [sbu],
                'Ruangan': [ruangan],
                'Tanggal': [tanggal],
                'Time In': [time_in],
                'Time Out': [time_out],
                'Status': ['booked']
            })

            bookings_df = pd.concat([bookings_df, new_booking], ignore_index=True)
            save_bookings(bookings_df)
            st.success('Booking confirmed!')

# Show existing bookings
st.subheader('Existing Bookings')
st.dataframe(bookings_df)
