﻿using Com.DanLiris.Service.Purchasing.Lib;
using Com.DanLiris.Service.Purchasing.Lib.Facades.Expedition;
using Com.DanLiris.Service.Purchasing.Lib.Services;
using Com.DanLiris.Service.Purchasing.Lib.Services.Expedition;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.IdentityModel.Tokens;
using Newtonsoft.Json.Serialization;
using System.Text;

namespace Com.DanLiris.Service.Purchasing.WebApi
{
    public class Startup
    {
        /* Hard Code */
        private string[] EXPOSED_HEADERS = new string[] { "Content-Disposition", "api-version", "content-length", "content-md5", "content-type", "date", "request-id", "response-time" };
        private string PURCHASING_POLICITY = "PurchasingPolicy";
        private string DEFAULT_CONNECTION = "DefaultConnection";
        private string SECRET = "Secret";
        
        public IConfiguration Configuration { get; }


        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        #region Register

        private void RegisterEndpoints()
        {

        }

        private void RegisterFacades(IServiceCollection services)
        {
            services
                .AddTransient<PurchasingDocumentExpeditionFacade>();
        }

        private void RegisterServices(IServiceCollection services)
        {
            services
                .AddTransient<PurchasingDocumentExpeditionService>()
                .AddScoped<IdentityService>()
                .AddScoped<HttpClientService>()
                .AddScoped<ValidateService>();
        }

        #endregion Register

        public void ConfigureServices(IServiceCollection services)
        {
            string connectionString = Configuration.GetConnectionString(DEFAULT_CONNECTION) ?? Configuration[DEFAULT_CONNECTION];

            /* Register */
            services.AddDbContext<PurchasingDbContext>(options => options.UseSqlServer(connectionString));
            RegisterEndpoints();
            RegisterFacades(services);
            RegisterServices(services);

            /* Versioning */
            services.AddApiVersioning(options => { options.DefaultApiVersion = new ApiVersion(1, 0); });

            /* Authentication */
            string Secret = Configuration.GetValue<string>(SECRET) ?? Configuration[SECRET];
            SymmetricSecurityKey Key = new SymmetricSecurityKey(Encoding.ASCII.GetBytes(Secret));

            services.AddAuthentication(JwtBearerDefaults.AuthenticationScheme)
                .AddJwtBearer(options =>
                {
                    options.TokenValidationParameters = new TokenValidationParameters
                    {
                        ValidateAudience = false,
                        ValidateIssuer = false,
                        ValidateLifetime = false,
                        IssuerSigningKey = Key
                    };
                });

            /* CORS */
            services.AddCors(options => options.AddPolicy(PURCHASING_POLICITY, builder =>
            {
                builder.AllowAnyOrigin()
                       .AllowAnyMethod()
                       .AllowAnyHeader()
                       .WithExposedHeaders(EXPOSED_HEADERS);
            }));

            /* API */
            services
               .AddMvcCore()
               .AddAuthorization()
               .AddJsonOptions(options => options.SerializerSettings.ContractResolver = new DefaultContractResolver())
               .AddJsonFormatters();
        }

        public void Configure(IApplicationBuilder app, IHostingEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }

            /* Update Database */
            using (var serviceScope = app.ApplicationServices.GetRequiredService<IServiceScopeFactory>().CreateScope())
            {
                PurchasingDbContext context = serviceScope.ServiceProvider.GetService<PurchasingDbContext>();
                context.Database.Migrate();
            }

            app.UseAuthentication();
            app.UseCors(PURCHASING_POLICITY);
            app.UseMvc();
        }
    }
}
