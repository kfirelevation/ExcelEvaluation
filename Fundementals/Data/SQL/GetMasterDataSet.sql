
SELECT        sales.order_items.order_id, sales.order_items.item_id,  production.products.product_name, production.products.model_year, production.brands.brand_name, sales.order_items.list_price, sales.order_items.quantity, sales.order_items.discount, 
                         sales.orders.order_date, (sales.staffs.first_name + ' ' + sales.staffs.last_name) as 'Sales Man', (sales.customers.first_name + ' ' + sales.customers.last_name) AS CutomerName, sales.customers.email As 'Customer email'
FROM            sales.order_items INNER JOIN
                         sales.orders ON sales.order_items.order_id = sales.orders.order_id INNER JOIN
                         production.products ON sales.order_items.product_id = production.products.product_id INNER JOIN
                         production.brands ON production.products.brand_id = production.brands.brand_id AND production.products.brand_id = production.brands.brand_id INNER JOIN
                         production.categories ON production.products.category_id = production.categories.category_id AND production.products.category_id = production.categories.category_id INNER JOIN
                         sales.customers ON sales.orders.customer_id = sales.customers.customer_id AND sales.orders.customer_id = sales.customers.customer_id AND sales.orders.customer_id = sales.customers.customer_id AND 
                         sales.orders.customer_id = sales.customers.customer_id INNER JOIN
                         sales.staffs ON sales.orders.staff_id = sales.staffs.staff_id AND sales.orders.staff_id = sales.staffs.staff_id AND sales.orders.staff_id = sales.staffs.staff_id AND sales.orders.staff_id = sales.staffs.staff_id
