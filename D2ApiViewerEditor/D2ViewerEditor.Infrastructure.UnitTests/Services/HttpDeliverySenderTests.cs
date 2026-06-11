using System.Net;
using D2ViewerEditor.Domain.Interfaces;
using D2ViewerEditor.Infrastructure.Services.Delivery;
using FluentAssertions;
using Microsoft.Extensions.Logging.Abstractions;
using NUnit.Framework;

namespace D2ViewerEditor.Infrastructure.UnitTests.Services;

/// <summary>
/// Wysyłka na returnUrl jako multipart/form-data (pole „file") + nagłówki idempotencji/hash.
/// </summary>
[TestFixture]
public class HttpDeliverySenderTests
{
    private static DeliveryDispatch Dispatch() =>
        new(Guid.NewGuid(), "https://recipient.example.com/inbox", new byte[] { 1, 2, 3, 4 }, "ABC123");

    [Test]
    public async Task SendAsync_PostsMultipartFormData_WithFileFieldAndHeaders()
    {
        var handler = new CapturingHandler(HttpStatusCode.OK);
        var sender = new HttpDeliverySender(new HttpClient(handler), NullLogger<HttpDeliverySender>.Instance);
        var dispatch = Dispatch();

        var result = await sender.SendAsync(dispatch);

        result.Outcome.Should().Be(DeliveryOutcome.Succeeded);
        handler.Method.Should().Be(HttpMethod.Post);
        handler.ContentType.Should().StartWith("multipart/form-data");
        // .NET zapisuje parametry Content-Disposition bez cudzysłowów: name=file; filename=document.docx
        handler.Body.Should().Contain("name=file");
        handler.Body.Should().Contain("filename=document.docx");
        handler.Body.Should().Contain("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        handler.RequestHeaders.Should().ContainKey("Idempotency-Key");
        handler.RequestHeaders["Idempotency-Key"].Should().Be(dispatch.DeliveryId.ToString());
        handler.RequestHeaders.Should().ContainKey("X-Content-SHA256");
        handler.RequestHeaders["X-Content-SHA256"].Should().Be("ABC123");
    }

    [Test]
    public async Task SendAsync_PermanentStatus_IsClassifiedPermanent()
    {
        var sender = new HttpDeliverySender(
            new HttpClient(new CapturingHandler(HttpStatusCode.UnprocessableEntity)),
            NullLogger<HttpDeliverySender>.Instance);

        var result = await sender.SendAsync(Dispatch());

        result.Outcome.Should().Be(DeliveryOutcome.PermanentError);
    }

    [Test]
    public async Task SendAsync_ServerError_IsRetryable()
    {
        var sender = new HttpDeliverySender(
            new HttpClient(new CapturingHandler(HttpStatusCode.ServiceUnavailable)),
            NullLogger<HttpDeliverySender>.Instance);

        var result = await sender.SendAsync(Dispatch());

        result.Outcome.Should().Be(DeliveryOutcome.RetryableError);
    }

    private sealed class CapturingHandler : HttpMessageHandler
    {
        private readonly HttpStatusCode _status;
        public HttpMethod? Method;
        public string? ContentType;
        public string? Body;
        public readonly Dictionary<string, string> RequestHeaders = new();

        public CapturingHandler(HttpStatusCode status) => _status = status;

        protected override async Task<HttpResponseMessage> SendAsync(
            HttpRequestMessage request, CancellationToken cancellationToken)
        {
            Method = request.Method;
            ContentType = request.Content?.Headers.ContentType?.ToString();
            // Odczyt zanim sender zwolni request/content (using w SendAsync).
            Body = request.Content is null ? null : await request.Content.ReadAsStringAsync(cancellationToken);
            foreach (var h in request.Headers)
                RequestHeaders[h.Key] = string.Join(",", h.Value);
            return new HttpResponseMessage(_status);
        }
    }
}
