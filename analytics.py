import pandas as pd
import matplotlib.pyplot as plt
import plotly.express as px
from sqlalchemy import create_engine
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import ColorScaleRule
import os
from datetime import datetime, timedelta
import seaborn as sns


class TicketSalesAnalyzer:
    def __init__(self):
        self.engine = create_engine('postgresql://postgres:776759@localhost:5432/Des1')
        plt.style.use('seaborn-v0_8')
        self.colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7', '#DDA0DD']
        print("[SUCCESS] Успешно подключились к базе данных через SQLAlchemy")

    def execute_query(self, query, description=""):
        """Выполнение SQL-запроса через SQLAlchemy"""
        try:
            df = pd.read_sql_query(query, self.engine)
            if description:
                print(f"[DATA] {description}: {len(df)} строк")
            return df
        except Exception as e:
            print(f"[ERROR] Ошибка выполнения запроса: {e}")
            return None

    def create_pie_chart(self):
        """Круговая диаграмма: распределение выручки по категориям"""
        query = """
        SELECT c.catname, SUM(s.pricepaid) as revenue
        FROM sale s
        JOIN events e ON s.eventid = e.eventid
        JOIN category c ON e.catid = c.catid
        GROUP BY c.catname
        ORDER BY revenue DESC;
        """

        df = self.execute_query(query, "Распределение выручки по категориям")
        if df is not None and len(df) > 0:
            plt.figure(figsize=(10, 8))
            plt.pie(df['revenue'], labels=df['catname'], autopct='%1.1f%%',
                    colors=self.colors, startangle=90)
            plt.title('Распределение выручки по категориям событий', fontsize=14, fontweight='bold')
            plt.tight_layout()
            os.makedirs('charts', exist_ok=True)
            plt.savefig('charts/pie_chart_revenue_by_category.png', dpi=300, bbox_inches='tight')
            plt.close()
            print("[SUCCESS] Создана круговая диаграмма: charts/pie_chart_revenue_by_category.png")

    def create_bar_chart(self):
        """Столбчатая диаграмма: топ площадок по событиям"""
        query = """
        SELECT v.venuename, COUNT(e.eventid) as event_count
        FROM venue v
        JOIN events e ON v.venueid = e.venueid
        GROUP BY v.venuename
        ORDER BY event_count DESC
        LIMIT 10;
        """

        df = self.execute_query(query, "Топ площадок по событиям")
        if df is not None and len(df) > 0:
            plt.figure(figsize=(12, 6))
            bars = plt.bar(range(len(df)), df['event_count'], color=self.colors[0])
            plt.title('Топ-10 площадок по количеству событий', fontsize=14, fontweight='bold')
            plt.xlabel('Площадки')
            plt.ylabel('Количество событий')
            plt.xticks(range(len(df)), df['venuename'], rotation=45, ha='right')

            for bar, count in zip(bars, df['event_count']):
                plt.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 5,
                         int(count), ha='center', va='bottom', fontsize=9)

            plt.tight_layout()
            plt.savefig('charts/bar_chart_top_venues.png', dpi=300, bbox_inches='tight')
            plt.close()
            print("[SUCCESS] Создана столбчатая диаграмма: charts/bar_chart_top_venues.png")

    def create_horizontal_bar_chart(self):
        """Горизонтальная столбчатая диаграмма: средний чек по штатам"""
        query = """
        SELECT u.state, 
               AVG(s.pricepaid) as avg_transaction,
               COUNT(s.saleid) as total_sales
        FROM "user" u
        JOIN sale s ON u.userid = s.buyerid
        GROUP BY u.state
        HAVING COUNT(s.saleid) > 100
        ORDER BY avg_transaction DESC
        LIMIT 15;
        """

        df = self.execute_query(query, "Средний чек по штатам")
        if df is not None and len(df) > 0:
            plt.figure(figsize=(12, 8))
            bars = plt.barh(range(len(df)), df['avg_transaction'], color=self.colors[1])
            plt.title('Средняя стоимость транзакции по штатам ($)', fontsize=14, fontweight='bold')
            plt.xlabel('Средняя стоимость транзакции ($)')
            plt.ylabel('Штаты')
            plt.yticks(range(len(df)), df['state'])

            for i, (bar, value) in enumerate(zip(bars, df['avg_transaction'])):
                plt.text(bar.get_width() + 1, bar.get_y() + bar.get_height() / 2,
                         f'${value:.2f}', va='center', fontsize=9)

            plt.tight_layout()
            plt.savefig('charts/horizontal_bar_avg_transaction.png', dpi=300, bbox_inches='tight')
            plt.close()
            print("[SUCCESS] Создана горизонтальная столбчатая диаграмма: charts/horizontal_bar_avg_transaction.png")

    def create_line_chart(self):
        """Линейный график: динамика продаж по месяцам"""
        query = """
        SELECT 
            EXTRACT(YEAR FROM s.saletime) as year,
            EXTRACT(MONTH FROM s.saletime) as month,
            COUNT(s.saleid) as total_sales,
            SUM(s.pricepaid) as total_revenue
        FROM sale s
        JOIN events e ON s.eventid = e.eventid
        JOIN venue v ON e.venueid = v.venueid
        GROUP BY year, month
        ORDER BY year, month;
        """

        df = self.execute_query(query, "Динамика продаж по месяцам")
        if df is not None and len(df) > 0:
            df['date'] = pd.to_datetime(
                df['year'].astype(int).astype(str) + '-' + df['month'].astype(int).astype(str) + '-01')

            fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(14, 10))

            ax1.plot(df['date'], df['total_sales'], marker='o', linewidth=2, color=self.colors[2])
            ax1.set_title('Динамика количества продаж по месяцам', fontsize=14, fontweight='bold')
            ax1.set_ylabel('Количество продаж')
            ax1.grid(True, alpha=0.3)

            ax2.plot(df['date'], df['total_revenue'], marker='s', linewidth=2, color=self.colors[3])
            ax2.set_title('Динамика выручки по месяцам', fontsize=14, fontweight='bold')
            ax2.set_ylabel('Выручка ($)')
            ax2.set_xlabel('Месяц')
            ax2.grid(True, alpha=0.3)

            plt.tight_layout()
            plt.savefig('charts/line_chart_sales_trends.png', dpi=300, bbox_inches='tight')
            plt.close()
            print("[SUCCESS] Создан линейный график: charts/line_chart_sales_trends.png")

    def create_histogram(self):
        """Гистограмма: распределение цен на билеты"""
        query = """
        SELECT l.priceperticket
        FROM listing l
        JOIN events e ON l.eventid = e.eventid
        JOIN category c ON e.catid = c.catid
        WHERE l.priceperticket BETWEEN 1 AND 500;
        """

        df = self.execute_query(query, "Распределение цен на билеты")
        if df is not None and len(df) > 0:
            plt.figure(figsize=(12, 6))
            plt.hist(df['priceperticket'], bins=30, color=self.colors[4], alpha=0.7, edgecolor='black')
            plt.title('Распределение цен на билеты', fontsize=14, fontweight='bold')
            plt.xlabel('Цена билета ($)')
            plt.ylabel('Количество билетов')
            plt.grid(True, alpha=0.3)

            plt.tight_layout()
            plt.savefig('charts/histogram_ticket_prices.png', dpi=300, bbox_inches='tight')
            plt.close()
            print("[SUCCESS] Создана гистограмма: charts/histogram_ticket_prices.png")

    def create_scatter_plot(self):
        """Точечная диаграмма: цена vs количество проданных билетов"""
        query = """
        SELECT 
            AVG(l.priceperticket) as avg_ticket_price,
            SUM(s.qtysold) as total_tickets_sold,
            c.catname
        FROM sale s
        JOIN listing l ON s.listid = l.listid
        JOIN events e ON s.eventid = e.eventid
        JOIN category c ON e.catid = c.catid
        GROUP BY c.catname
        HAVING SUM(s.qtysold) > 100;
        """

        df = self.execute_query(query, "Цена vs количество проданных билетов")
        if df is not None and len(df) > 0:
            plt.figure(figsize=(12, 8))
            scatter = plt.scatter(df['avg_ticket_price'], df['total_tickets_sold'],
                                  c=df['avg_ticket_price'], cmap='viridis', s=100, alpha=0.6)

            plt.colorbar(scatter, label='Средняя цена билета ($)')
            plt.title('Связь между ценой билета и количеством проданных билетов', fontsize=14, fontweight='bold')
            plt.xlabel('Средняя цена билета ($)')
            plt.ylabel('Общее количество проданных билетов')
            plt.grid(True, alpha=0.3)

            for i, row in df.iterrows():
                plt.annotate(row['catname'],
                             (row['avg_ticket_price'], row['total_tickets_sold']),
                             xytext=(5, 5), textcoords='offset points', fontsize=8)

            plt.tight_layout()
            plt.savefig('charts/scatter_price_vs_quantity.png', dpi=300, bbox_inches='tight')
            plt.close()
            print("[SUCCESS] Создана точечная диаграмма: charts/scatter_price_vs_quantity.png")

    def create_interactive_slider_chart(self):
        """Интерактивный график с временным слайдером"""
        query = """
        SELECT 
            DATE(s.saletime) as sale_date,
            c.catname,
            AVG(s.pricepaid) as avg_price,
            COUNT(s.saleid) as daily_sales,
            SUM(s.qtysold) as daily_tickets
        FROM sale s
        JOIN events e ON s.eventid = e.eventid
        JOIN category c ON e.catid = c.catid
        WHERE s.saletime >= '2008-01-01'
        GROUP BY sale_date, c.catname
        ORDER BY sale_date;
        """

        df = self.execute_query(query, "Данные для интерактивного графика")
        if df is not None and len(df) > 0:
            df['sale_date'] = pd.to_datetime(df['sale_date'])
            df['month_year'] = df['sale_date'].dt.to_period('M').astype(str)

            fig = px.scatter(df,
                             x="daily_sales",
                             y="avg_price",
                             size="daily_tickets",
                             color="catname",
                             animation_frame="month_year",
                             hover_name="catname",
                             title="Интерактивная динамика продаж по категориям",
                             labels={"daily_sales": "Ежедневные продажи",
                                     "avg_price": "Средняя цена ($)"})

            fig.write_html("charts/interactive_sales_chart.html")
            print("[SUCCESS] Создан интерактивный график: charts/interactive_sales_chart.html")

    def create_interactive_category_sales(self):
        """Интерактивная диаграмма: динамика продаж по категориям с анимацией"""
        query = """
        SELECT 
            DATE(s.saletime) as sale_date,
            EXTRACT(YEAR FROM s.saletime) as year,
            EXTRACT(MONTH FROM s.saletime) as month,
            c.catname,
            COUNT(s.saleid) as daily_sales,
            SUM(s.pricepaid) as daily_revenue,
            SUM(s.qtysold) as daily_tickets
        FROM sale s
        JOIN events e ON s.eventid = e.eventid
        JOIN category c ON e.catid = c.catid
        WHERE s.saletime BETWEEN '2008-01-01' AND '2008-12-31'
        GROUP BY sale_date, c.catname, year, month
        ORDER BY sale_date, c.catname;
        """

        df = self.execute_query(query, "Динамика продаж по категориям для интерактивного графика")

        if df is not None and len(df) > 0:
            monthly_data = df.groupby(['year', 'month', 'catname']).agg({
                'daily_sales': 'sum',
                'daily_revenue': 'sum',
                'daily_tickets': 'sum'
            }).reset_index()

            fig = px.bar(
                monthly_data,
                x="catname",
                y="daily_revenue",
                color="catname",
                animation_frame="month",
                animation_group="catname",
                range_y=[0, monthly_data['daily_revenue'].max() * 1.1],
                title="Интерактивная динамика продаж по категориям (2008 год)",
                labels={
                    "daily_revenue": "Выручка ($)",
                    "catname": "Категория событий",
                    "month": "Месяц"
                },
                hover_data=["daily_sales", "daily_tickets"]
            )

            fig.layout.updatemenus[0].buttons[0].args[1]["frame"]["duration"] = 1000
            fig.layout.updatemenus[0].buttons[0].args[1]["transition"]["duration"] = 500

            fig.write_html("charts/interactive_category_sales.html")
            print(
                "[SUCCESS] Создана интерактивная диаграмма продаж по категориям: charts/interactive_category_sales.html")

            return True
        return False

    def create_advanced_interactive_dashboard(self):
        """Продвинутая интерактивная панель с несколькими графиками"""
        query = """
        SELECT 
            DATE(s.saletime) as sale_date,
            EXTRACT(MONTH FROM s.saletime) as month,
            c.catname,
            u.state,
            AVG(s.pricepaid) as avg_price,
            COUNT(s.saleid) as sales_count,
            SUM(s.pricepaid) as total_revenue,
            SUM(s.qtysold) as total_tickets
        FROM sale s
        JOIN events e ON s.eventid = e.eventid
        JOIN category c ON e.catid = c.catid
        JOIN "user" u ON s.buyerid = u.userid
        WHERE s.saletime BETWEEN '2008-01-01' AND '2008-12-31'
        GROUP BY sale_date, c.catname, u.state, month
        HAVING SUM(s.pricepaid) > 0
        ORDER BY sale_date;
        """

        df = self.execute_query(query, "Данные для продвинутой панели")

        if df is not None and len(df) > 0:
            fig = px.scatter(
                df,
                x="sales_count",
                y="total_revenue",
                size="total_tickets",
                color="catname",
                hover_name="state",
                animation_frame="month",
                size_max=60,
                title="Динамика продаж: количество vs выручка по штатам",
                labels={
                    "sales_count": "Количество продаж",
                    "total_revenue": "Общая выручка ($)",
                    "catname": "Категория",
                    "state": "Штат"
                }
            )

            fig.write_html("charts/advanced_sales_dashboard.html")
            print("[SUCCESS] Создана продвинутая интерактивная панель: charts/advanced_sales_dashboard.html")

            return True
        return False

    def export_to_excel(self, dataframes_dict, filename):
        """Экспорт данных в Excel с форматированием"""
        try:
            os.makedirs('exports', exist_ok=True)
            filepath = f'exports/{filename}'

            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                for sheet_name, df in dataframes_dict.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

                    workbook = writer.book
                    worksheet = writer.sheets[sheet_name]

                    worksheet.freeze_panes = "A2"

                    worksheet.auto_filter.ref = worksheet.dimensions

                    numeric_columns = df.select_dtypes(include=['number']).columns
                    for col_idx, col_name in enumerate(numeric_columns, 1):
                        col_letter = chr(64 + col_idx)
                        range_str = f"{col_letter}2:{col_letter}{len(df) + 1}"

                        rule = ColorScaleRule(
                            start_type="min", start_color="FFAA0000",
                            mid_type="percentile", mid_value=50, mid_color="FFFFFF00",
                            end_type="max", end_color="FF00AA00"
                        )
                        worksheet.conditional_formatting.add(range_str, rule)

            total_sheets = len(dataframes_dict)
            total_rows = sum(len(df) for df in dataframes_dict.values())

            print(f"[SUCCESS] Создан файл {filename}, {total_sheets} листов, {total_rows} строк")
            return True

        except Exception as e:
            print(f"[ERROR] Ошибка при экспорте в Excel: {e}")
            return False

    def prepare_data_for_excel_export(self):
        """Подготовка данных для экспорта в Excel"""
        queries = {
            "Sales_Summary": """
                SELECT c.catname,
                       COUNT(s.saleid) as total_sales,
                       SUM(s.qtysold) as total_tickets,
                       SUM(s.pricepaid) as total_revenue,
                       AVG(s.pricepaid) as avg_sale_amount
                FROM sale s
                JOIN events e ON s.eventid = e.eventid
                JOIN category c ON e.catid = c.catid
                GROUP BY c.catname
                ORDER BY total_revenue DESC;
            """,
            "User_Geography": """
                SELECT city, state,
                       COUNT(*) as user_count
                FROM "user"
                GROUP BY city, state
                HAVING COUNT(*) > 10
                ORDER BY user_count DESC;
            """,
            "Venue_Performance": """
                SELECT v.venuename, v.venuecity, v.venuestate,
                       COUNT(DISTINCT e.eventid) as total_events,
                       SUM(s.pricepaid) as total_revenue,
                       AVG(s.pricepaid) as avg_revenue_per_event
                FROM venue v
                JOIN events e ON v.venueid = e.venueid
                JOIN sale s ON e.eventid = s.eventid
                GROUP BY v.venueid, v.venuename, v.venuecity, v.venuestate
                ORDER BY total_revenue DESC;
            """
        }

        dataframes = {}
        for sheet_name, query in queries.items():
            df = self.execute_query(query, f"Подготовка данных для {sheet_name}")
            if df is not None:
                dataframes[sheet_name] = df

        return dataframes

    def run_complete_analysis(self):
        """Запуск полного анализа"""
        print("=" * 50)
        print("ЗАПУСК ПОЛНОГО АНАЛИЗА TICKET SALES")
        print("=" * 50)

        self.create_pie_chart()
        self.create_bar_chart()
        self.create_horizontal_bar_chart()
        self.create_line_chart()
        self.create_histogram()
        self.create_scatter_plot()

        self.create_interactive_slider_chart()
        self.create_interactive_category_sales()
        self.create_advanced_interactive_dashboard()

        excel_data = self.prepare_data_for_excel_export()
        self.export_to_excel(excel_data, "ticket_sales_analysis.xlsx")

        print("\n[SUCCESS] АНАЛИЗ ЗАВЕРШЕН!")
        print("[INFO] Результаты сохранены в папках: charts/, exports/")


def main():
    try:
        analyzer = TicketSalesAnalyzer()
        analyzer.run_complete_analysis()
    except Exception as e:
        print(f"[ERROR] Произошла ошибка: {e}")


if __name__ == "__main__":
    main()