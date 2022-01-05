import subprocess

def asrun(ascript):
	osa = subprocess.Popen(['osascript', '-'], stdin = subprocess.PIPE,
		stdout = subprocess.PIPE, universal_newlines=True)
	return osa.communicate(ascript)[0]

class codeBuilder:
	def build_code(self):
		code = '\n'.join(self.header + self.footer)
		self.last_code = code
		return code 

	def add_snippet(self,head,foot=''):
		self.header.append(head)
		if len(foot) >0: self.footer.insert(0,foot)

	def __repr__(self):
		return self.myself

	def __str__(self):
		#list of children
		ret = []
		for key,value in self.properties.items():
			ret.append('{}: {}'.format(key,value))
		return '\n'.join(ret)

	def __init__(self, myself, parent = None, process_name='Word'):
		self.process_name = process_name
		self.myself = myself
		self.parent = parent
		self.children = []
		self.properties = {}
		self.myselfstring = self.getMySelfStr()
		# print (self.myselfstring)
		self.reset_code()

	def getMySelfStr(self):
		myselfstring = self.myself
		if self.parent:
			myselfstring = "{} of {}".format(myselfstring, self.parent.getMySelfStr())
		return myselfstring

	def reset_code(self):
		self.header = []
		self.footer = []
		self.bring_process()

	def bring_process(self):
		self.add_snippet(
			'''tell application "System Events" to tell process "{}"'''.format(self.process_name),

			'end tell'
			)

	def details(self):
		if len(self.properties): return

		self.add_snippet('properties of {}'.format(self.myselfstring))
		ret = asrun(self.build_code()).split(':')
		self.header.pop()
		# properties parsing
		key = ret[0]
		for x in ret[1:]:
			xsp = x.split(',')
			value = ', '.join([ y.strip() for y in xsp[:-1] ]).strip()

			exc_prop = [
			'minimum value'
			'orientation',
			'position',
			'size',
			'role',
			'subrole'
			]

			if key not in exc_prop and value != "missing value" and value !="":
				# print('{}: {}'.format(key, value))
				self.properties[key] = value

			key = xsp[-1].strip()

	def updateName(self):
		# print(self.myselfstring)
		self.details()
		self.myself = '{} "{}"'.format(self.properties['class'], self.properties['name'])
		self.myselfstring = self.getMySelfStr()


	def getChildren(self):
		self.add_snippet('ui elements of {}'.format(self.myselfstring))
		ret = asrun(self.build_code())
		self.header.pop()
		# ui children parsing
		uis = [r.strip() for r in [r.split('of')[0] for r in ret.split(',')]]
		return uis

	def add_child(self, ui):
		self.children.append( codeBuilder(ui, self, self.process_name) )
		return self.children[-1]

	def build_tree(self,depth = 1):
		if len(self.children)==0:
			if depth == 0: return
			if depth <0: depth = -1

			for n,ui in enumerate(self.getChildren()):
				if ui=='':
					continue

				if ui.split()[-1].isnumeric():
					self.add_child(ui).build_tree( depth - 1 )
				else:
					ch = self.add_child('ui element {}'.format(n+1))
					ch.updateName()
					ch.build_tree( depth - 1 )


		return len(self.children)

	def run(self):
		ret = asrun(self.build_code())
		return ret


# below example is to navigate UI hierarchy of ribbon tab area of MS Office
if __name__ == '__main__':
	C = codeBuilder('tab group 1 of window 1')
	
	while 1:
		print()
		print(C.myself)
		print('---------------------------------------')
		C.details()
		print(C)
		print('---------------------------------------')

		hasChildren = C.build_tree(1)
		if hasChildren:
			for i, child in enumerate(C.children):
				print(i, child.myself)
		print('b: go to parent')

		inp = input()
		if inp.isnumeric():
			no = int(inp)
			C = C.children[int(no)]
		else:
			if inp == 'b':
				if C.parent: C = C.parent
			if inp == 'q':
				break
			if inp == 'c':
				print(C.last_code)
			if inp == 'p':
				print(C.myselfstring)

