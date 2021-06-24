from O365 import Account
from O365 import MSGraphProtocol
import datetime as dt
import time

protocol_graph = MSGraphProtocol()

scopes = protocol_graph.get_scopes_for('calendar')

credentials = ('ae5d19ab-4ab2-4d70-8f52-7bcbfb1274f3', 'CzR9D9XOK3L-~ub07CVfQl_i1nPs~6jp~4')

account = Account(credentials, auth_flow_type='authorization', tenant_id='e964c47e-9b02-444c-890c-9593c775ab82')

if not account.authenticate(scopes=scopes):
    print("Não autenticou. :(")
    quit()

# print('Emails\n')
# mailbox = account.mailbox("rodrigo.ramos@lambda3.com.br")

# inbox = mailbox.inbox_folder()

# for message in inbox.get_messages():
#     print(message)

print("Calendário")

while True:
    print("\nObtendo eventos")
    schedule = account.schedule()

    calendar = schedule.get_default_calendar()
    today = dt.date.today()

    start_query = today
    end_query = today + dt.timedelta(weeks=1)

    query = calendar.new_query('start').greater_equal(start_query)
    query.chain("and").on_attribute('end').less_equal(end_query)

    print("Eventos de {f} até {t}".format(f=start_query, t=end_query))
    events = calendar.get_events(query=query, include_recurring=True)

    print("Escrendo arquivo .pal... ")
    with open("/home/rramos/.pal/lambda3.pal", "w+") as f:
        f.write("L3 Lambda3\n")

        for event in events:
            print("evento", end='')
            print(event)

            if event.start.date().isoformat() == today.isoformat() and event.start.hour < dt.datetime.now().hour:
                continue

            f.write(event.start.strftime("%Y%m%d"))
            f.write(" ")
            f.write(event.start.strftime("%H:%M"))
            f.write(" - ")
            f.write(event.end.strftime("%H:%M"))
            f.write(" ")
            f.write(event.subject)
            f.write(" (")
            f.write(event.organizer.name)
            f.write(")")
            f.write("\n")

    print("Feito!\nAguardando próxima iteração.")
    time.sleep(300) # 5 mins

