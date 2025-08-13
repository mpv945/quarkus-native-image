package com.ewancle;

/*import com.ewancle.entity.Fruit;
import io.quarkus.panache.common.Sort;*/
import io.quarkus.virtual.threads.VirtualThreads;
import io.smallrye.mutiny.Uni;
import jakarta.inject.Inject;
import jakarta.ws.rs.GET;
import jakarta.ws.rs.Path;
import jakarta.ws.rs.Produces;
import jakarta.ws.rs.core.MediaType;

import java.util.List;
import java.util.concurrent.ExecutorService;

@Path("/hello")
// @RunOnVirtualThread 和响应式二选一
public class ExampleResource {

    @Inject
    @VirtualThreads
    ExecutorService vThreads;

    @GET
    @Produces(MediaType.TEXT_PLAIN)
    public String hello() {
        vThreads.execute(this::findAll);
        return Thread.currentThread().getName()+"Hello from Quarkus REST";
    }

    void findAll() {
        System.out.println(Thread.currentThread().getName()+"Hello from Quarkus REST --> vThreads");
    }


    /*@GET
    public Uni<List<Fruit>> get() {
        return Fruit.listAll(Sort.by("name"));
    }*/
}
