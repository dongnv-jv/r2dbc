package com.example.r2dbc.service;

import com.example.r2dbc.entity.Employee;
import com.example.r2dbc.repository.EmployeeRepository;
import lombok.RequiredArgsConstructor;
import org.springframework.data.r2dbc.core.DatabaseClient;
import org.springframework.data.r2dbc.core.R2dbcEntityTemplate;
import org.springframework.stereotype.Service;
import reactor.core.publisher.Flux;
import reactor.core.publisher.Mono;

@Service
@RequiredArgsConstructor
public class EmployeeService {

    private final EmployeeRepository employeeRepository;
    private final R2dbcEntityTemplate r2dbcEntityTemplate;
    private final DatabaseClient databaseClient;

    public Flux<Employee> saveAll() {
        Flux<Employee> studentFlux = Flux.range(1, 1_000_000)
                .map(i -> new Employee("Dong "+i, "Student-" + i));
        return employeeRepository.saveAll(studentFlux);

    }
    public Flux<Employee> getAll() {
        return r2dbcEntityTemplate.select(Employee.class).all();
    }
    public Flux<Employee> get() {
        return databaseClient
                .execute("select * from employee")
                .map((row, rowMetadata) -> new Employee(row.get("name", String.class),row.get("department", String.class))
                )
                .all();
    }
    public Mono<String> callStoredProcedure(int inputParam) {
        return databaseClient
                .execute("CALL stored_procedure_name(:inputParam, @outputParam)")
                .bind("inputParam", inputParam)
                .map((row, rowMetadata) -> row.get("outputParam", String.class))
                .one();
    }
}
