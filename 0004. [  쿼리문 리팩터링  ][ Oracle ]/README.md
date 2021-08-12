# [  쿼리문 리팩터링  ][ Oracle ] <br><br>

- 설명 : <br>

     경영실적 분석 데이터를 산출하는 쿼리문이다. <br>
     똑같은 SELECT절을 반복해서 서브쿼리로 사용해서 중복이 발생했다. <br>
     그래서 이부분을 with문으로 캐싱하여 재사용 했다. <br>
     
     ```diff
     ※ 보안과 쿼리문의 이해를 돕기위해 쿼리문은 추상화 하여 작성.
     ```
     
     




<br><br><br><br>





### 1. 최종 데이터 형태. <br>

![image](https://user-images.githubusercontent.com/54452587/129153854-2b724d60-3baf-4603-afb8-a201e9dd9724.png)

최종 데이터의 형태를 큰 흐름을 분석해보면 아래와 같다. <br>

매출액 - 매출원가 = 매출총이익 <br>
매출총이익 - 일반관리비 = 영업이익 <br>
영업이익 + 영업외수익 - 영업외비용 = 경상이익 <br>



<br><br><br><br>



### 2. 기초자료 with문으로 구성하기. <br>

탭 기능의 경우 현재 열린 탭이 있으면 해당 탭으로 이동하고 . <br>


```sql

매출_WITH AS(
      SELECT *
        FROM 매출_TABLE
       WHERE 년도 = '2021'
)

                                                                                                         
 , 비용_WITH AS(
      SELECT *
        FROM 비용_TABLE
       WHERE 년도 = '2021'
)

```
<br><br><br><br>



### 3. 기초자료를 사용해서 2차자료를 with문으로 구성하기. <br>

 . <br>


```sql


, 제품_WITH AS(
      SELECT *
        FROM 매출_WITH
       WHERE 매출종류 = '제품'
)


, 상품_WITH AS(
      SELECT *
        FROM 매출_WITH
       WHERE 매출종류 = '상품'
)


, 재료비_WITH AS (
      SELECT *
        FROM 비용_WITH
       WHERE 비용종류 = '재료비'
)

                                                                                                         
 , 노무비_WITH AS(
      SELECT *
        FROM 비용_WITH
       WHERE 비용종류 = '노무비'
)


, 제조경비_WITH AS(
      SELECT *
        FROM 비용_WITH
       WHERE 비용종류 = '제조경비'
)


, 일반관리비_WITH AS(
      SELECT *
        FROM 비용_WITH
       WHERE 비용종류 = '일반관리비'
)


, 영업외수익_WITH AS(
      SELECT *
        FROM 매출_WITH
       WHERE 매출종류 = '영업외수익'
)


, 일반관리비_WITH AS(
      SELECT *
        FROM 비용_WITH
       WHERE 비용종류 = '일반관리비'
)



```
<br><br><br><br>





### 5. 2차자료를 사용해서 3차자료를 WITH문으로 구성하기. <br>

 . <br>


```sql

 , 매출총이익_WITH AS(
      SELECT *
        FROM (
                SELECT '매출' VALGBN, MAEMON , ODDGBN , 금액
                  FROM MECHUL_TABLE_WITH 
                 UNION ALL
                SELECT '매출' VALGBN, MAEMON , ODDGBN , (-금액) AMOUNT
                  FROM SALES_COST_WITH
             )
       WHERE 년도 = '2021'
)


 , 영업이익_WITH AS(
      SELECT *
        FROM 비용_TABLE
       WHERE 년도 = '2021'
)


 , 경상이익_WITH AS(
      SELECT *
        FROM 비용_TABLE
       WHERE 년도 = '2021'
)

```
<br><br><br><br>





### 4. 최종 데이터 구성하기. <br>

 . <br>


```sql

매출_WITH AS(
      SELECT *
        FROM 매출_TABLE
       WHERE 년도 = '2021'
)

                                                                                                         
 , 비용_WITH AS(
      SELECT *
        FROM 비용_TABLE
       WHERE 년도 = '2021'
)

```
<br><br><br><br>















<br><br><br><br>




```java
 
```









