import extract_msg

file = r'M:\automate311\Requête Gatineau_ca.msg'  # Replace with yours
msg = extract_msg.Message(file)
msg_sender = msg.sender
msg_date = msg.date
msg_subj = msg.subject
msg_message = msg.body
msg.close()

text = 'Body: {}'.format(msg_message)
# email = 'Sender: {}'.format(msg_sender))

lines = text.splitlines()


dic = {}

for line in lines[5:25:2] + ['\n'.join(lines[27:-2])] + [lines[-1]]:
	k, v = line.split(' : ', 1)
	dic[k] = v

for i, k in dic.items():
	print(i, k)
