﻿using AutoMapper;
using Com.DanLiris.Service.Purchasing.Lib.Interfaces;
using Com.DanLiris.Service.Purchasing.Lib.Models.GarmentInternalPurchaseOrderModel;
using Com.DanLiris.Service.Purchasing.Lib.Services;
using Com.DanLiris.Service.Purchasing.Lib.ViewModels.GarmentInternalPurchaseOrderViewModel;
using Com.DanLiris.Service.Purchasing.Test.Helpers;
using Com.DanLiris.Service.Purchasing.WebApi.Controllers.v1.GarmentInternalPurchaseOrderControllers;
using Com.Moonlay.NetCore.Lib.Service;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Moq;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Net;
using System.Security.Claims;
using System.Text;
using Xunit;

namespace Com.DanLiris.Service.Purchasing.Test.Controllers.GarmentInternalPurchaseOrderControllerTests
{
    public class GarmentInternalPurchaseOrderControllerTest
    {
        private GarmentInternalPurchaseOrderViewModel ViewModel
        {
            get
            {
                return new GarmentInternalPurchaseOrderViewModel
                {
                    Items = new List<GarmentInternalPurchaseOrderItemViewModel>
                    {
                        new GarmentInternalPurchaseOrderItemViewModel()
                    }
                };
            }
        }

        private GarmentInternalPurchaseOrder Model
        {
            get
            {
                return new GarmentInternalPurchaseOrder { };
            }
        }

        private ServiceValidationExeption GetServiceValidationExeption()
        {
            Mock<IServiceProvider> serviceProvider = new Mock<IServiceProvider>();
            List<ValidationResult> validationResults = new List<ValidationResult>();
            System.ComponentModel.DataAnnotations.ValidationContext validationContext = new System.ComponentModel.DataAnnotations.ValidationContext(ViewModel, serviceProvider.Object, null);
            return new ServiceValidationExeption(validationContext, validationResults);
        }

        private Mock<IServiceProvider> GetServiceProvider()
        {
            var serviceProvider = new Mock<IServiceProvider>();
            serviceProvider
                .Setup(x => x.GetService(typeof(IdentityService)))
                .Returns(new IdentityService() { Token = "Token", Username = "Test" });

            serviceProvider
                .Setup(x => x.GetService(typeof(IHttpClientService)))
                .Returns(new HttpClientTestService());

            return serviceProvider;
        }

        private GarmentInternalPurchaseOrderController GetController(Mock<IGarmentInternalPurchaseOrderFacade> facadeM, Mock<IValidateService> validateM, Mock<IMapper> mapper)
        {
            var user = new Mock<ClaimsPrincipal>();
            var claims = new Claim[]
            {
                new Claim("username", "unittestusername")
            };
            user.Setup(u => u.Claims).Returns(claims);

            var servicePMock = GetServiceProvider();
            if(validateM != null)
            {
                servicePMock
                    .Setup(x => x.GetService(typeof(IValidateService)))
                    .Returns(validateM.Object);
            }

            var controller = new GarmentInternalPurchaseOrderController(servicePMock.Object, mapper.Object, facadeM.Object)
            {
                ControllerContext = new ControllerContext()
                {
                    HttpContext = new DefaultHttpContext()
                    {
                        User = user.Object
                    }
                }
            };
            controller.ControllerContext.HttpContext.Request.Headers["Authorization"] = "Bearer unittesttoken";
            controller.ControllerContext.HttpContext.Request.Path = new PathString("/v1/unit-test");
            controller.ControllerContext.HttpContext.Request.Headers["x-timezone-offset"] = "7";

            return controller;
        }

        protected int GetStatusCode(IActionResult response)
        {
            return (int)response.GetType().GetProperty("StatusCode").GetValue(response, null);
        }

        [Fact]
        public void Should_Success_Create_Multiple_Data()
        {
            var validateMock = new Mock<IValidateService>();
            validateMock.Setup(s => s.Validate(It.IsAny<GarmentInternalPurchaseOrderViewModel>())).Verifiable();

            var mockMapper = new Mock<IMapper>();
            mockMapper.Setup(x => x.Map<List<GarmentInternalPurchaseOrder>>(It.IsAny<List<GarmentInternalPurchaseOrderViewModel>>()))
                .Returns(new List<GarmentInternalPurchaseOrder>());

            var mockFacade = new Mock<IGarmentInternalPurchaseOrderFacade>();
            mockFacade.Setup(x => x.CreateMultiple(It.IsAny<List<GarmentInternalPurchaseOrder>>(), "unittestusername", 7))
               .ReturnsAsync(1);

            var controller = GetController(mockFacade, validateMock, mockMapper);

            var response = controller.Post(new List<GarmentInternalPurchaseOrderViewModel> { ViewModel }).Result;
            Assert.Equal((int)HttpStatusCode.Created, GetStatusCode(response));
        }

        [Fact]
        public void Should_Validate_Create_Multiple_Data()
        {
            var validateMock = new Mock<IValidateService>();
            validateMock.Setup(s => s.Validate(It.IsAny<GarmentInternalPurchaseOrderViewModel>())).Throws(GetServiceValidationExeption());

            var mockMapper = new Mock<IMapper>();
            var mockFacade = new Mock<IGarmentInternalPurchaseOrderFacade>();

            var controller = GetController(mockFacade, validateMock, mockMapper);

            var response = controller.Post(new List<GarmentInternalPurchaseOrderViewModel> { ViewModel }).Result;
            Assert.Equal((int)HttpStatusCode.BadRequest, GetStatusCode(response));
        }

        [Fact]
        public void Should_Error_Create_Data()
        {
            var mockMapper = new Mock<IMapper>();
            var mockFacade = new Mock<IGarmentInternalPurchaseOrderFacade>();

            var controller = new GarmentInternalPurchaseOrderController(GetServiceProvider().Object, mockMapper.Object, mockFacade.Object);

            var response = controller.Post(new List<GarmentInternalPurchaseOrderViewModel>()).Result;
            Assert.Equal((int)HttpStatusCode.InternalServerError, GetStatusCode(response));
        }

        [Fact]
        public void Should_Success_Get_All_Data_By_User()
        {
            var mockFacade = new Mock<IGarmentInternalPurchaseOrderFacade>();

            mockFacade.Setup(x => x.Read(It.IsAny<int>(), It.IsAny<int>(), It.IsAny<string>(), null, It.IsAny<string>()))
                .Returns(Tuple.Create(new List<GarmentInternalPurchaseOrder>(), 0, new Dictionary<string, string>()));

            var mockMapper = new Mock<IMapper>();
            mockMapper.Setup(x => x.Map<List<GarmentInternalPurchaseOrderViewModel>>(It.IsAny<List<GarmentInternalPurchaseOrder>>()))
                .Returns(new List<GarmentInternalPurchaseOrderViewModel> { ViewModel });

            GarmentInternalPurchaseOrderController controller = GetController(mockFacade, null, mockMapper);
            var response = controller.GetByUser();
            Assert.Equal((int)HttpStatusCode.OK, GetStatusCode(response));
        }

        [Fact]
        public void Should_Success_Get_All_Data_By_User_With_Filter()
        {
            var mockFacade = new Mock<IGarmentInternalPurchaseOrderFacade>();

            mockFacade.Setup(x => x.Read(It.IsAny<int>(), It.IsAny<int>(), It.IsAny<string>(), null, It.IsAny<string>()))
                .Returns(Tuple.Create(new List<GarmentInternalPurchaseOrder>(), 0, new Dictionary<string, string>()));

            var mockMapper = new Mock<IMapper>();
            mockMapper.Setup(x => x.Map<List<GarmentInternalPurchaseOrderViewModel>>(It.IsAny<List<GarmentInternalPurchaseOrder>>()))
                .Returns(new List<GarmentInternalPurchaseOrderViewModel> { ViewModel });

            GarmentInternalPurchaseOrderController controller = GetController(mockFacade, null, mockMapper);
            var response = controller.GetByUser(filter:"{ 'IsPosted': false }");
            Assert.Equal((int)HttpStatusCode.OK, GetStatusCode(response));
        }

        [Fact]
        public void Should_Error_Get_All_Data()
        {
            var mockFacade = new Mock<IGarmentInternalPurchaseOrderFacade>();

            mockFacade.Setup(x => x.Read(It.IsAny<int>(), It.IsAny<int>(), It.IsAny<string>(), null, It.IsAny<string>()))
                .Returns(Tuple.Create(new List<GarmentInternalPurchaseOrder>(), 0, new Dictionary<string, string>()));

            var mockMapper = new Mock<IMapper>();
            mockMapper.Setup(x => x.Map<List<GarmentInternalPurchaseOrderViewModel>>(It.IsAny<List<GarmentInternalPurchaseOrder>>()))
                .Returns(new List<GarmentInternalPurchaseOrderViewModel> { ViewModel });

            GarmentInternalPurchaseOrderController controller = new GarmentInternalPurchaseOrderController(GetServiceProvider().Object, mockMapper.Object, mockFacade.Object);
            var response = controller.Get();
            Assert.Equal((int)HttpStatusCode.InternalServerError, GetStatusCode(response));
        }

        [Fact]
        public void Should_Error_Get_All_Data_By_User()
        {
            var mockFacade = new Mock<IGarmentInternalPurchaseOrderFacade>();

            mockFacade.Setup(x => x.Read(It.IsAny<int>(), It.IsAny<int>(), It.IsAny<string>(), null, It.IsAny<string>()))
                .Returns(Tuple.Create(new List<GarmentInternalPurchaseOrder>(), 0, new Dictionary<string, string>()));

            var mockMapper = new Mock<IMapper>();
            mockMapper.Setup(x => x.Map<List<GarmentInternalPurchaseOrderViewModel>>(It.IsAny<List<GarmentInternalPurchaseOrder>>()))
                .Returns(new List<GarmentInternalPurchaseOrderViewModel> { ViewModel });

            GarmentInternalPurchaseOrderController controller = new GarmentInternalPurchaseOrderController(GetServiceProvider().Object, mockMapper.Object, mockFacade.Object);
            var response = controller.GetByUser();
            Assert.Equal((int)HttpStatusCode.InternalServerError, GetStatusCode(response));
        }

        [Fact]
        public void Should_Success_Get_Data_By_Id()
        {
            var mockFacade = new Mock<IGarmentInternalPurchaseOrderFacade>();

            mockFacade.Setup(x => x.ReadById(It.IsAny<int>()))
                .Returns(new GarmentInternalPurchaseOrder());
            mockFacade.Setup(x => x.CheckDuplicate(It.IsAny<GarmentInternalPurchaseOrder>()))
                .Returns(true);

            var mockMapper = new Mock<IMapper>();
            mockMapper.Setup(x => x.Map<GarmentInternalPurchaseOrderViewModel>(It.IsAny<GarmentInternalPurchaseOrder>()))
                .Returns(ViewModel);

            GarmentInternalPurchaseOrderController controller = new GarmentInternalPurchaseOrderController(GetServiceProvider().Object, mockMapper.Object, mockFacade.Object);
            var response = controller.Get(It.IsAny<int>());
            Assert.Equal((int)HttpStatusCode.OK, GetStatusCode(response));
        }

        [Fact]
        public void Should_Error_Get_Data_By_Id()
        {
            var mockFacade = new Mock<IGarmentInternalPurchaseOrderFacade>();

            mockFacade.Setup(x => x.ReadById(It.IsAny<int>()))
                .Returns(new GarmentInternalPurchaseOrder());

            var mockMapper = new Mock<IMapper>();

            GarmentInternalPurchaseOrderController controller = new GarmentInternalPurchaseOrderController(GetServiceProvider().Object, mockMapper.Object, mockFacade.Object);
            var response = controller.Get(It.IsAny<int>());
            Assert.Equal((int)HttpStatusCode.InternalServerError, GetStatusCode(response));
        }

        [Fact]
        public void Should_Success_Split_Data()
        {
            var validateMock = new Mock<IValidateService>();
            validateMock.Setup(s => s.Validate(It.IsAny<GarmentInternalPurchaseOrderViewModel>())).Verifiable();

            var mockMapper = new Mock<IMapper>();
            mockMapper.Setup(x => x.Map<GarmentInternalPurchaseOrder>(It.IsAny<GarmentInternalPurchaseOrderViewModel>()))
                .Returns(new GarmentInternalPurchaseOrder());

            var mockFacade = new Mock<IGarmentInternalPurchaseOrderFacade>();
            mockFacade.Setup(x => x.Split(It.IsAny<int>(), It.IsAny<GarmentInternalPurchaseOrder>(), "unittestusername", 7))
               .ReturnsAsync(1);

            var controller = GetController(mockFacade, validateMock, mockMapper);

            var response = controller.Put(It.IsAny<int>(), It.IsAny<GarmentInternalPurchaseOrderViewModel>()).Result;
            Assert.Equal((int)HttpStatusCode.NoContent, GetStatusCode(response));
        }

        [Fact]
        public void Should_Validate_Split_Data()
        {
            var validateMock = new Mock<IValidateService>();
            validateMock.Setup(s => s.Validate(It.IsAny<GarmentInternalPurchaseOrderViewModel>())).Throws(GetServiceValidationExeption());

            var mockMapper = new Mock<IMapper>();
            var mockFacade = new Mock<IGarmentInternalPurchaseOrderFacade>();

            var controller = GetController(mockFacade, validateMock, mockMapper);

            var response = controller.Put(It.IsAny<int>(), It.IsAny<GarmentInternalPurchaseOrderViewModel>()).Result;
            Assert.Equal((int)HttpStatusCode.BadRequest, GetStatusCode(response));
        }

        [Fact]
        public void Should_Error_Split_Data()
        {
            var mockMapper = new Mock<IMapper>();
            var mockFacade = new Mock<IGarmentInternalPurchaseOrderFacade>();

            var controller = new GarmentInternalPurchaseOrderController(GetServiceProvider().Object, mockMapper.Object, mockFacade.Object);

            var response = controller.Put(It.IsAny<int>(), It.IsAny<GarmentInternalPurchaseOrderViewModel>()).Result;
            Assert.Equal((int)HttpStatusCode.InternalServerError, GetStatusCode(response));
        }

        [Fact]
        public void Should_Success_Delete_Data()
        {
            var validateMock = new Mock<IValidateService>();
            var mockMapper = new Mock<IMapper>();

            var mockFacade = new Mock<IGarmentInternalPurchaseOrderFacade>();
            mockFacade.Setup(x => x.Delete(It.IsAny<int>(), "unittestusername"))
               .ReturnsAsync(1);

            var controller = GetController(mockFacade, validateMock, mockMapper);

            var response = controller.Delete(It.IsAny<int>()).Result;
            Assert.Equal((int)HttpStatusCode.NoContent, GetStatusCode(response));
        }

        [Fact]
        public void Should_Error_Delete_Data()
        {
            var mockMapper = new Mock<IMapper>();
            var mockFacade = new Mock<IGarmentInternalPurchaseOrderFacade>();

            var controller = new GarmentInternalPurchaseOrderController(GetServiceProvider().Object, mockMapper.Object, mockFacade.Object);

            var response = controller.Delete(It.IsAny<int>()).Result;
            Assert.Equal((int)HttpStatusCode.InternalServerError, GetStatusCode(response));
        }

        [Fact]
        public void Should_Error_Get_All_Data_By_Tags()
        {
            var mockMapper = new Mock<IMapper>();
            var mockFacade = new Mock<IGarmentInternalPurchaseOrderFacade>();

            GarmentInternalPurchaseOrderController controller = new GarmentInternalPurchaseOrderController(GetServiceProvider().Object, mockMapper.Object, mockFacade.Object);
            var response = controller.GetByTags(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>(), It.IsAny<string>());
            Assert.Equal((int)HttpStatusCode.InternalServerError, GetStatusCode(response));
        }

        [Fact]
        public void Should_Success_Get_All_Data_By_Tags_With_ShipmentDate()
        {
            var mockFacade = new Mock<IGarmentInternalPurchaseOrderFacade>();

            mockFacade.Setup(x => x.ReadByTags(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<DateTimeOffset>(), It.IsAny<DateTimeOffset>(), It.IsAny<string>()))
                .Returns(new List<GarmentInternalPurchaseOrder>());

            var mockMapper = new Mock<IMapper>();
            mockMapper.Setup(x => x.Map<List<GarmentInternalPurchaseOrderViewModel>>(It.IsAny<List<GarmentInternalPurchaseOrder>>()))
                .Returns(new List<GarmentInternalPurchaseOrderViewModel>());

            GarmentInternalPurchaseOrderController controller = GetController(mockFacade, null, mockMapper);
            var response = controller.GetByTags(null, null, "2018-10-31", "2018-10-31");
            Assert.Equal((int)HttpStatusCode.OK, GetStatusCode(response));
        }

		[Fact]
		public void Should_Success_Get_All_Data_ByName()
		{
			var mockFacade = new Mock<IGarmentInternalPurchaseOrderFacade>();

			mockFacade.Setup(x => x.ReadName(null,"{}"))
				.Returns(new List<GarmentInternalPurchaseOrder>());

			var mockMapper = new Mock<IMapper>();
			mockMapper.Setup(x => x.Map<List<GarmentInternalPurchaseOrderViewModel>>(It.IsAny<List<GarmentInternalPurchaseOrder>>()))
				.Returns(new List<GarmentInternalPurchaseOrderViewModel>());

			GarmentInternalPurchaseOrderController controller = GetController(mockFacade, null, mockMapper);
			var response = controller.ByName(null, null);
			Assert.Equal((int)HttpStatusCode.OK, GetStatusCode(response));
		}
        #region Report Purchase Order Internal
        [Fact]
        public void Should_Success_Get_Report_PO()
        {
            var mockFacade = new Mock<IGarmentInternalPurchaseOrderFacade>();

            mockFacade.Setup(x => x.GetReportPO(It.IsAny<DateTime>(), It.IsAny<DateTime>(), It.IsAny<int>(), It.IsAny<int>(), It.IsAny<string>(), It.IsAny<int>(), It.IsAny<string>()))
                .Returns(Tuple.Create(new List<GarmentInternalPurchaseOrderReportViewModel>(), 5));

            var mockMapper = new Mock<IMapper>();
            GarmentInternalPurchaseOrderController controller = GetController(mockFacade, null, mockMapper);
            var response = controller.GetReportPO(It.IsAny<DateTime>(), It.IsAny<DateTime>(), It.IsAny<int>(), It.IsAny<int>(), It.IsAny<string>());
            Assert.Equal((int)HttpStatusCode.OK, GetStatusCode(response));
        }

        [Fact]
        public void Should_Success_Get_Xls_PO()
        {
            var mockFacade = new Mock<IGarmentInternalPurchaseOrderFacade>();
            mockFacade.Setup(x => x.GenerateExcelPO(It.IsAny<DateTime>(), It.IsAny<DateTime>(), It.IsAny<int>(), It.IsAny<string>()))
                .Returns(new MemoryStream());

            var mockMapper = new Mock<IMapper>();

            var user = new Mock<ClaimsPrincipal>();
            var claims = new Claim[]
            {
                new Claim("username", "unittestusername")
            };
            user.Setup(u => u.Claims).Returns(claims);
            GarmentInternalPurchaseOrderController controller = new GarmentInternalPurchaseOrderController(GetServiceProvider().Object, mockMapper.Object, mockFacade.Object);
            controller.ControllerContext = new ControllerContext()
            {
                HttpContext = new DefaultHttpContext()
                {
                    User = user.Object
                }
            };

            controller.ControllerContext.HttpContext.Request.Headers["x-timezone-offset"] = "0";
            var response = controller.GetXlsPO(It.IsAny<DateTime>(), It.IsAny<DateTime>());
            Assert.Equal("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", response.GetType().GetProperty("ContentType").GetValue(response, null));

        }

        [Fact]
        public void Should_Error_Get_Report_PO()
        {
            var mockFacade = new Mock<IGarmentInternalPurchaseOrderFacade>();

            var mockMapper = new Mock<IMapper>();


            GarmentInternalPurchaseOrderController controller = GetController(mockFacade, null, mockMapper);
            var response = controller.GetReportPO(It.IsAny<DateTime>(), It.IsAny<DateTime>(),It.IsAny<int>(), It.IsAny<int>(), It.IsAny<string>());
            Assert.Equal((int)HttpStatusCode.InternalServerError, GetStatusCode(response));
        }

        [Fact]
        public void Should_Error_Get_Xls_PO()
        {
            var mockFacade = new Mock<IGarmentInternalPurchaseOrderFacade>();

            var mockMapper = new Mock<IMapper>();

            GarmentInternalPurchaseOrderController controller = new GarmentInternalPurchaseOrderController(GetServiceProvider().Object, mockMapper.Object, mockFacade.Object);

            var response = controller.GetXlsPO(It.IsAny<DateTime>(), It.IsAny<DateTime>());
            Assert.Equal((int)HttpStatusCode.InternalServerError, GetStatusCode(response));

        }
        #endregion
    }
}
