# 🏭 Polymer Production Scheduler

A Streamlit application for optimizing polymer production scheduling with constraint programming.

## Features

- 📊 **Data Analysis**: Process plant capacities, inventory constraints, and demand forecasts
- ⚡ **Optimization**: Use Google OR-Tools for constraint programming optimization
- 📈 **Visualization**: Interactive charts for production schedules and inventory levels
- 💾 **Reporting**: Generate detailed Excel reports with production schedules

## Deployment on Streamlit Cloud

1. Create a new GitHub repository and upload all these files
2. Go to [Streamlit Cloud](https://streamlit.io/cloud)
3. Click "New app" and connect your GitHub repository
4. Set the main file path to `app.py`
5. Deploy!

## Local Development

```bash
# Install dependencies
pip install -r requirements.txt

# Run the app
streamlit run app.py