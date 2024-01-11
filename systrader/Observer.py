from abc import ABC, abstractmethod
import pickle

# observer pattern
# 데이터 대 전략 구조 구현을 위한 코드로, '1대 다' mapping를 지원
# Controller: 한 데이터가 변경되면, 해당 데이터를 팔로우 하던 observer들이 작동함
# 아직, Controller는 동기적으로 처리됨

class Observer(ABC):
    @abstractmethod
    def update(self, data):
        pass

class Controller:
    def __init__(self):
        self._observers = []
        self._date = None
        self._name = None

    def add_observer(self, observer):
        self._observers.append(observer)

    def remove_observer(self, observer):
        self._observers.remove(observer)

    def set_name(self, name):
        self._name = name

    def set_data(self, new_data):
        if new_data != self._date:
            self._data = new_data
            self.nofity_observers()

    def nofity_observers(self):
        for observer in self._observers:
            observer.update(self._data)

    def save(self):
        path = './Controller/' + self.name + '.pkl'
        with open(path, 'wb') as file:
            pickle.dump(self, file)

class SignalObserver(Observer):
    def update(self, data):
        # 1. signal class로 전달 / signal.update()
        # 2. database 저장
        pass

############################################################
# 사용예시
# controller = Controller()
# stratege1 = SignalObserver()
# stratege2 = SignalObserver()
#
# controller.add_observer(stratege1)
# controller.add_observer(stratege2)
#
# controller.set_data('데이터1')
# controller.set_data('데이터2')
############################################################