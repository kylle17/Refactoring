# [  쿼리문 리팩터링  ][ Oracle ] <br><br>



```diff
     ※ 보안과 쿼리문의 이해를 돕기위해 쿼리문은 추상화 하여 작성.
```

- 설명 : <br>

     경영실적 분석 데이터를 산출하는 쿼리문이다. <br>
     기존 쿼리문은 똑같은 SELECT절을 반복해서 서브쿼리로 사용하여 중복이 발생했다. <br>
     그래서 이부분을 WITH문으로 캐싱하여 재사용해서 중복을 없앴다. <br>
     

     
     




<br><br><br><br>





### 1. 최종 데이터 형태. <br>

![image](https://user-images.githubusercontent.com/54452587/129153854-2b724d60-3baf-4603-afb8-a201e9dd9724.png)

최종 데이터의 형태를 큰 흐름을 분석해보면 아래와 같다. <br>

매출액 - 매출원가 = 매출총이익 <br>
매출총이익 - 일반관리비 = 영업이익 <br>
영업이익 + 영업외수익 - 영업외비용 = 경상이익 <br>



<br><br><br><br>



### 2. 기초자료를 WITH문으로 구성하기. <br>

 필요한 모든 데이터를 DB에서 가져와서 WITH문으로 구성한다. <br>
 WHERE절을 사용해서 필요한 데이터만 가져온다. <br>


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



### 3. 기초자료를 사용해서 2차자료를 WITH문으로 구성하기. <br>

기초자료 WITH문에서 필요한 데이터만 가져와서 <br>
계획과 실적으로 구분하고 GROUP BY하여 합계금액을 구한다. <br>


```sql


, 제품_WITH AS(
      SELECT SUM(DECODE( 구분 , 계획 , 금액 , 0 )) 금액1 , SUM(  구분 , 실적 , 금액 , 0  )) 금액2
        FROM 매출_WITH
       WHERE 매출종류 = '제품'
       GROUP BY 1
)


, 상품_WITH AS(
      SELECT SUM(DECODE( 구분 , 계획 , 금액 , 0 )) 금액1 , SUM(  구분 , 실적 , 금액 , 0  )) 금액2
        FROM 매출_WITH
       WHERE 매출종류 = '상품' 
       GROUP BY 1
)


, 재료비_WITH AS(
      SELECT SUM(DECODE( 구분 , 계획 , 금액 , 0 )) 금액1 , SUM(  구분 , 실적 , 금액 , 0  )) 금액2
        FROM 비용_WITH
       WHERE 비용종류 = '재료비'  
       GROUP BY 1
)        


, 노무비_WITH AS(
      SELECT SUM(DECODE( 구분 , 계획 , 금액 , 0 )) 금액1 , SUM(  구분 , 실적 , 금액 , 0  )) 금액2
        FROM 비용_WITH
       WHERE 비용종류 = '노무비'
       GROUP BY 1
)           


, 제조경비_WITH AS(
      SELECT SUM(DECODE( 구분 , 계획 , 금액 , 0 )) 금액1 , SUM(  구분 , 실적 , 금액 , 0  )) 금액2
        FROM 비용_WITH
       WHERE 비용종류 = '제조경비' 
       GROUP BY 1
) 
  
  
, 일반관리비_WITH AS(
      SELECT SUM(DECODE( 구분 , 계획 , 금액 , 0 )) 금액1 , SUM(  구분 , 실적 , 금액 , 0  )) 금액2
        FROM 비용_WITH
       WHERE 비용종류 = '일반관리비'  
)  


, 영업외수익_WITH AS(
      SELECT SUM(DECODE( 구분 , 계획 , 금액 , 0 )) 금액1 , SUM(  구분 , 실적 , 금액 , 0  )) 금액2
        FROM 매출_WITH
       WHERE 매출종류 = '영업외수익' 
       GROUP BY 1
)  


, 영업외비용_WITH AS(
      SELECT SUM(DECODE( 구분 , 계획 , 금액 , 0 )) 금액1 , SUM(  구분 , 실적 , 금액 , 0  )) 금액2
        FROM 비용_WITH
       WHERE 비용종류 = '영업외비용'  
)  




```
<br><br><br><br>







### 4. 2차자료를 사용해서 3차자료를 WITH문으로 구성하기.  <br>

 . <br>


```sql

, 매출액_WITH AS(
      SELECT SUM(DECODE( 구분 , 계획 , 금액 , 0 )) 금액1 , SUM(  구분 , 실적 , 금액 , 0  )) 금액2
        FROM (
               SELECT 금액1 , 금액2 FROM 제품_WITH
                UNION ALL
               SELECT 금액1 , 금액2 FROM 상품_WITH
             )
       GROUP BY 1
)
 
 
, 매출원가_WITH AS(
      SELECT SUM(DECODE( 구분 , 계획 , 금액 , 0 )) 금액1 , SUM(  구분 , 실적 , 금액 , 0  )) 금액2
        FROM ( 
               SELECT * FROM 재료비_WITH 
                UNION ALL
               SELECT * FROM 노무비_WITH 
                UNION ALL
               SELECT * FROM 제조경비_WITH 
              )
       GROUP BY 1
)


, 매출총이익_WITH AS(
      SELECT SUM(DECODE( 구분 , 계획 , 금액 , 0 )) 금액1 , SUM(  구분 , 실적 , 금액 , 0  )) 금액2
        FROM (
               SELECT 금액1 , 금액2 FROM 매출액_WITH
                UNION ALL
               SELECT (- 금액1) , (- 금액2) FROM 매출원가_WITH 
             )
       GROUP BY 1     
)  
        

, 영업이익_WITH AS(
      SELECT SUM(DECODE( 구분 , 계획 , 금액 , 0 )) 금액1 , SUM(  구분 , 실적 , 금액 , 0  )) 금액2
        FROM (
               SELECT 금액1 , 금액2 FROM 매출총이익_WITH
                UNION ALL
               SELECT (- 금액1) , (- 금액2) FROM 일반관리비_WITH 
             )
       GROUP BY 1   
)          
     

, 경상이익_WITH AS(
      SELECT SUM(DECODE( 구분 , 계획 , 금액 , 0 )) 금액1 , SUM(  구분 , 실적 , 금액 , 0  )) 금액2
        FROM (
               SELECT 금액1 , 금액2 FROM 영업이익_WITH
                UNION ALL
               SELECT 금액1 , 금액2 FROM 영업외수익_WITH 
                UNION ALL
               SELECT (- 금액1) , (- 금액2) FROM 영업외비용_WITH                
             )
       GROUP BY 1 
)  

        




```
<br><br><br><br>





### 5. 위에서 구성한 2차,3차자료를 사용해서 최종 데이터 구성하기. <br>

 . <br>


```sql
      SELECT ' + 매출액' title, 금액1 계획 , 금액2 실적
        FROM 매출액_WITH  
       UNION ALL       
      SELECT '     제품' title , 금액1 계획 , 금액2 실적
        FROM 제품_WITH
       UNION ALL
      SELECT '     상품' title , 금액1 계획 , 금액2 실적
        FROM 상품_WITH
        
       UNION ALL      
      SELECT ' - 매출원가' title , 금액1 계획 , 금액2 실적
        FROM 매출원가_WITH 
       UNION ALL       
      SELECT '     재료비' title, 금액1 계획 , 금액2 실적
        FROM 재료비_WITH	
       UNION ALL  
      SELECT '     노무비' title, 금액1 계획 , 금액2 실적
        FROM 노무비_WITH 
       UNION ALL    
      SELECT '     제조경비' title, 금액1 계획 , 금액2 실적
        FROM 제조경비_WITH 
        
       UNION ALL
      SELECT ' = 매출 총이익' title, 금액1 계획 , 금액2 실적
        FROM 매출총이익_WITH
       UNION ALL    
      SELECT ' - 일반 관리비' title, 금액1 계획 , 금액2 실적
        FROM 일반관리비_WITH
       UNION ALL        
      SELECT ' = 영업 이익' title, 금액1 계획 , 금액2 실적
        FROM 영업이익_WITH        
       UNION ALL          
      SELECT ' + 영업외 수익' title, 금액1 계획 , 금액2 실적
        FROM 영업외수익_WITH
       UNION ALL        
      SELECT ' - 영업외 비용' title, 금액1 계획 , 금액2 실적
        FROM 영업외비용_WITH
       UNION ALL        
      SELECT ' = 경상 이익' title, 금액1 계획 , 금액2 실적
        FROM 경상이익_WITH

```
<br><br><br><br>




### 마무리 <br>

쿼리문 실행 결과를 확인해보면 아래와 같다. <br><br>


수정 전 <br>
![image](https://user-images.githubusercontent.com/54452587/129995224-288d8f79-c395-4a0a-8fb5-9655ec9aa06d.png)<br><br>

수정 후 <br>
![image](https://user-images.githubusercontent.com/54452587/129995251-bc0e2995-ae21-42c0-b6d1-68b3385751f9.png)<br><br>


실행속도가 훨씬 빨라진 것을 확인할 수 있다. <br>
WITH문을 단계별로 구성해서 보기에도 편하고 필요한 부분만 찾아볼 수 있다. <br>

<br><br><br><br>





