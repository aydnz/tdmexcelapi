public class Startup
{
    public void ConfigureServices(IServiceCollection services)
    {
        // CORS'u etkinleştiriyoruz
        services.AddCors(options =>
        {
            options.AddPolicy("AllowAll", builder =>
            {
                builder.AllowAnyOrigin()  // Flutter uygulamanızın çalıştığı adresi belirtin (örneğin 'http://localhost:3000')
                       .AllowAnyMethod()
                       .AllowAnyHeader();
            });
        });

        services.AddControllers();
    }

    public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
    {
        if (env.IsDevelopment())
        {
            app.UseDeveloperExceptionPage();
        }

        // CORS middleware'ini buraya ekliyoruz
        app.UseCors("AllowAll");

        app.UseRouting();

        app.UseEndpoints(endpoints =>
        {
            endpoints.MapControllers();
        });
    }
}
